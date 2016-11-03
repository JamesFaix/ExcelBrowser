using System.Diagnostics;
using System.Threading;
using NUnit.Framework;
using xlApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelBrowser.Interop.Test {

    [TestFixture]
    class ExcelAppConverterTest {

        #region Infrastructure

        private int StartProcess(string name = "Excel") {
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
        public void ExcelAppConverter_CanGetAppFromProcess() {
          
            var sw = new Stopwatch();
            var timeoutMs = 10000;

            var id = StartProcess();
            try {
                var process = Process.GetProcessById(id);
                Assert.NotNull(process);

                xlApp app = null;

                //Wait for app to start
                sw.Start();
                while(sw.ElapsedMilliseconds < timeoutMs && app == null) {
                    app = process.AsExcelApp();
                }

                Assert.NotNull(app);
            }
            finally {
                KillProcess(id);
            }
        }

        [Test]
        public void ExcelAppConverter_CanGetProcessFromApp() {         
            var app = new xlApp();
            try {
                var process = app.AsProcess();
                Assert.NotNull(process);
            }
            finally {
                app.Quit();
            }
        }
    }
}
