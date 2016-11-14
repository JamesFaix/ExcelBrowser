using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using NUnit.Framework;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using System.Threading.Tasks;
using System.Threading;

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

        private int StartExcelWithWorkbook() {
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) +
                @"/../../TestWorkbook.xlsx"; //Move up through /bin/Debug

            using (var process = Process.Start("EXCEL", path)) {
                return process.Id;
            }
        }

        #endregion

        [Test]
        public void ExcelAppConverter_CanGetAppFromProcess() {

            var id = StartProcess();
            try {
                var process = Process.GetProcessById(id);
                Assert.NotNull(process);

                xlApp app = process.AsExcelApp();

                Assert.NotNull(app);
            }
            finally {
                KillProcess(id);
            }
        }

        [Test]
        public void ExcelAppConverter_CanGetAppFromMainWindowHandleIfWorkbookIsOpen() {
            var processId = StartExcelWithWorkbook();
            try {
                var process = Process.GetProcessById(processId);
                Assert.NotNull(process);

                var fetcher = new Fetcher<int>(
                    getValue: () => process.MainWindowHandle.ToInt32(),
                    valueFilter: x => (x != 0) && ExcelAppConverter.GetClassNameFromWindowHandle(x) == "XLMAIN",
                    timeoutSeconds: 5);

                var handle = fetcher.Fetch();
                Assert.AreNotEqual(0, handle);

                var app = ExcelAppConverter.AppFromMainWindowHandle(handle);
                Assert.NotNull(app);
            }
            finally {
                KillProcess(processId);
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

        [Test]
        public void ExcelAppConverter_CanGetAppFromMainWindowHandleIfNoWorkbooksOpen() {
            var processId = StartProcess();
            try {
                var process = Process.GetProcessById(processId);
                Assert.NotNull(process);

                var fetcher = new Fetcher<int>(
                    getValue: () => process.MainWindowHandle.ToInt32(),
                    valueFilter: x => x != 0,
                    timeoutSeconds: 5);

                var handle = fetcher.Fetch();
                Assert.AreNotEqual(0, handle);

                var app = ExcelAppConverter.AppFromMainWindowHandle(handle);
                Assert.NotNull(app);
            }
            finally {
                KillProcess(processId);
            }
        }


        [Test]
        public void PrintProcessMainWindowHandles() {
            var id = StartExcelWithWorkbook();
            var process = Process.GetProcessById(id);
            try {
                Task.Run(() => {
                    while (true) {
                        foreach (var handle in process.WindowHandles()) {
                            var className = ExcelAppConverter.GetClassNameFromWindowHandle(handle.ToInt32());
                            Debug.WriteLine($"Process: {id}, Hwnd: {handle}, ClassName: {className}");
                        }
                        Thread.Sleep(100);
                    }
                });
                Thread.Sleep(10000);
            }
            finally {
                KillProcess(id);
            }
        }
    }
}
