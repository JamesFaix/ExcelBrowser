using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBrowser.Processes;
using System.Diagnostics;
using System.Threading;

namespace ExcelBrowser.Processes.Test {

    [TestFixture]
    public class ProcessMonitorTest {

        private int StartProcess() {
            using (var process = Process.Start("Notepad")) {
                return process.Id;
            }
        }

        private void KillProcess(int id) {
            using (var process = Process.GetProcessById(id)) {
                process.Kill();
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
        public void ProcessMonitor_CanStartProcess() {

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
        public void ProcessMonitor_WillPublishChanges() {

            var pm = new ProcessMonitor();
            pm.RefreshMilliseconds = 10;

            //Values event will fill
            var triggered = false;
            var started = new int[0];
            var stopped = new int[0];

            //Update values on process changes
            pm.ProcessChange += (sender, e) => {
                started = e.StartedProcessIds;
                stopped = e.StoppedProcessIds;
                triggered = true;
            };

            var id = StartProcess();
            try {
                var timeoutMs = 1000;
                var sw = new Stopwatch();
                sw.Start();
                while (sw.ElapsedMilliseconds < timeoutMs) {
                    //Wait until event fires
                    if (!triggered) {
                        Thread.Sleep(pm.RefreshMilliseconds);
                    }
                    else {
                        Assert.AreEqual(id, started.Single());
                        Assert.AreEqual(0, stopped.Count());
                    }
                }
            }
            finally {
                KillProcess(id);
            }
        }
    }
}
