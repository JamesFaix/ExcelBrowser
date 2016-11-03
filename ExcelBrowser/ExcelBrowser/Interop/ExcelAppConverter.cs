using System;
using System.Diagnostics;
using xlApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelBrowser.Interop {

    public static class ExcelAppConverter {
        
        /// <summary>
        /// Gets the Windows Process associated with the given Excel instance.
        /// </summary>
        /// <param name="app">The application.</param>
        public static Process AsProcess(this xlApp app) {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var mainWindowHandle = app.Hwnd;
            var processId = WindowsApiUtil.GetProcessIdFromWindowHandle(mainWindowHandle);
            return Process.GetProcessById(processId);
        }
        
        /// <summary>
        /// Gets the Excel instance running in the given process, or null if none exists.
        /// </summary>
        /// <param name="process">The process.</param>
        public static xlApp AsExcelApp(this Process process) {
            if (process == null) throw new ArgumentNullException(nameof(process));
            
            return AppFromMainWindowHandle(process.MainWindowHandle.ToInt32());
        }

        /// <summary>
        /// Gets the Excel instance running in the process that has the given ID, or null if none exists.
        /// </summary>
        /// <param name="processId">The process identifier.</param>
        public static xlApp AppFromProcessId(int processId) => 
            Process.GetProcessById(processId)?.AsExcelApp();

        /// <summary>
        /// Gets the Excel instance whose main window has the given handle, or null if none exists.
        /// </summary>
        /// <param name="handle">The handle.</param>
        public static xlApp AppFromMainWindowHandle(int handle) =>
            WindowsApiUtil.ExcelApplicationFromMainWindowHandle(handle);
    }
}
