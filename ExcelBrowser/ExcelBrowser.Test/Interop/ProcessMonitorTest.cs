using System.Diagnostics;
using System.Linq;
using System.Threading;
using NUnit.Framework;

namespace ExcelBrowser.Interop.Test {

    [TestFixture]
    public class ProcessMonitorTest {

        #region Infrastructure

        private int StartProcess(string name = "Notepad") {
            using (var process = Process.Start(name)) {
                return process.Id;
            }
        }

        //Kill in a finally block!
        private void KillProcess(int id) {
            using (var process = Process.GetProcessById(id)) {
                process.Kill();
            }
        }

        #endregion

        [Test]
        public void Process_CanStartProcess() {
            /* This test does not test the ProcessMonitor class itself.
             * It tests the assumption that we can easily start processes, 
             * which will be required to implement further tests. 
             */

            //Check running processes
            var firstCheck = Process.GetProcesses().Select(p => p.Id).ToArray();

            //Start a process
            var newId = StartProcess();
            //Thread.Sleep(10); //If OS is heavily loaded, a timeout may be necessary

            //Check again
            var secondCheck = Process.GetProcesses().Select(p => p.Id).ToArray();

            try {
                //New process should be only difference
                var difference = secondCheck.Except(firstCheck).Single();
                Assert.AreEqual(newId, difference);
            }
            finally {
                KillProcess(newId);
            }
        }

        [Test]
        public void ProcessMonitor_DefaultRefreshIs500() {
            var pm = new ProcessMonitor();
            Assert.AreEqual(500, pm.RefreshMilliseconds);
        }

        [Test]
        public void ProcessMonitor_CanSetRefresh() {
            var pm = new ProcessMonitor();
            pm.RefreshMilliseconds = 20;
            Assert.AreEqual(20, pm.RefreshMilliseconds);
        }
        
        [Test]
        public void ProcessMonitor_WillNotPublishChangesWhenNoProcessChanges() {
            var pm = new ProcessMonitor();
            pm.RefreshMilliseconds = 10;
            Thread.Sleep(pm.RefreshMilliseconds); //Wait for one refresh so the monitor picks up any existing processes.

            //Test timeout settings
            var timeoutMs = 1000; //More than refresh milliseconds!
            var sw = new Stopwatch();

            pm.ProcessChange += (sender, e) => Assert.Fail("Event should not be fired.");

            //Wait until timeout
            sw.Start();
            while (sw.ElapsedMilliseconds < timeoutMs) {
                Thread.Sleep(pm.RefreshMilliseconds);
            }
        }

        [Test]
        public void ProcessMonitor_WillNotPublishChangesBeforeRefresh() {
            var pm = new ProcessMonitor();
            Thread.Sleep(pm.RefreshMilliseconds); //Wait for one refresh so the monitor picks up any existing processes.

            //Test timeout settings
            var timeoutMs = 250; //Less than refresh milliseconds!
            var sw = new Stopwatch();

            pm.ProcessChange += (sender, e) => Assert.Fail("Event should not be fired.");

            //Start a process
            var id = StartProcess();
            try {
                sw.Start();
                //Wait until timeout
                while (sw.ElapsedMilliseconds < timeoutMs) {
                    Thread.Sleep(pm.RefreshMilliseconds);
                }
            }
            finally {
                KillProcess(id);
            }
        }

        [Test]
        public void ProcessMonitor_WillPublishChangesAfterOneOrMoreRefreshes() {
            var pm = new ProcessMonitor();
            pm.RefreshMilliseconds = 10;
            Thread.Sleep(pm.RefreshMilliseconds); //Wait for one refresh so the monitor picks up any existing processes.

            //Test timeout settings
            var timeoutMs = 5000;
            var sw = new Stopwatch();

            //Update values on process changes
            int[] started = null;
            int[] stopped = null;

            pm.ProcessChange += (sender, e) => {
                started = e.StartedProcessIds;
                stopped = e.StoppedProcessIds;
            };

            //Start a process
            var id = StartProcess();
            try {
                sw.Start();
                while (sw.ElapsedMilliseconds < timeoutMs) {
                    //Wait until event fires
                    if (started == null) {
                        Thread.Sleep(pm.RefreshMilliseconds);
                    }
                    else {
                        //Only one process should have been added, and none removed
                        Assert.AreEqual(id, started.Single());
                        Assert.AreEqual(0, stopped.Count());
                    }
                }
                //Event should have fired at least once.
                Assert.IsNotNull(started);
            }
            finally {
                KillProcess(id);
            }
        }

        [Test]
        public void ProcessMonitor_NameFilterWillHideProcessesThatReturnFalse() {
            var pm = new ProcessMonitor();
            pm.RefreshMilliseconds = 10;
            pm.NameFilter = (name => name.ToUpper() == "NOTEPAD");
            Thread.Sleep(pm.RefreshMilliseconds); //Wait for one refresh so the monitor picks up any existing processes.

            //Test timeout settings
            var timeoutMs = 5000;
            var sw = new Stopwatch();
            
            pm.ProcessChange += (sender, e) => Assert.Fail("Event should not be fired.");

            //Start a process
            var id = StartProcess("Calc");
            try {
                //Wait until timeout;
                sw.Start();
                while (sw.ElapsedMilliseconds < timeoutMs) {
                    Thread.Sleep(pm.RefreshMilliseconds);
                }
            }
            finally {
                KillProcess(id);
            }
        }

        [Test]
        public void ProcessMonitor_NameFilterWillNotHideProcessesThatReturnTrue() {
            var pm = new ProcessMonitor();
            pm.RefreshMilliseconds = 10;
            pm.NameFilter = (name => name.ToUpper() == "NOTEPAD");
            Thread.Sleep(pm.RefreshMilliseconds); //Wait for one refresh so the monitor picks up any existing processes.

            //Test timeout settings
            var timeoutMs = 5000;
            var sw = new Stopwatch();

            //Update values on process changes
            int[] started = null;
            int[] stopped = null;

            pm.ProcessChange += (sender, e) => {
                started = e.StartedProcessIds;
                stopped = e.StoppedProcessIds;
            };

            //Start a process
            var id = StartProcess("Notepad");
            try {
                sw.Start();
                while (sw.ElapsedMilliseconds < timeoutMs) {
                    //Wait until event fires
                    if (started == null) {
                        Thread.Sleep(pm.RefreshMilliseconds);
                    }
                    else {
                        //Only one process should have been added, and none removed
                        Assert.AreEqual(id, started.Single());
                        Assert.AreEqual(0, stopped.Count());
                    }
                }
                //Event should have fired at least once.
                Assert.IsNotNull(started);
            }
            finally {
                KillProcess(id);
            }
        }
    }
}
