using System.Diagnostics;
using ExcelBrowser.Interop;
using ExcelBrowser.Model;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Controller {

    public class SessionCommands {

        public void ActivateApp(AppId id) {
            var app = GetApp(id.ProcessId);
            app.Activate();
        }

        public void ActivateBook(BookId id, int windowIndex = 1) => 
            ActivateWindow(new WindowId(id.ProcessId, id.BookName, windowIndex));

        public void ActivateWindow(WindowId id) {
            var app = GetApp(id.ProcessId);
            app.Activate();
            
            var win = GetWindow(id);
            win.Activate();
        }

        public void ActivateSheet(SheetId id, int windowIndex = 1) {
            var app = GetApp(id.ProcessId);
            app.Activate();

            var win = GetWindow(new WindowId(id.ProcessId, id.BookName, windowIndex));
            win.Activate();

            var sheet = GetSheet(id);
            sheet.Select();
        }
        
        //public void SetWindowState(WindowId id, WindowState state) {
        //    try {
        //        var app = AppFactory.FromProcessId(id.ProcessId);
        //        Debug.Assert(app != null);
        //    }
        //    catch (Exception x) {
        //        //TODO: Use catch block
        //    }
        //}

        private Application GetApp(int processId) {
            var app= AppFactory.FromProcessId(processId);
            Debug.Assert(app != null);
            return app;
        }

        private Workbook GetBook(int processId, string bookName) {
            var app = GetApp(processId);
            var book = app.Workbooks[bookName];
            Debug.Assert(book != null);
            return book;
        }

        private Window GetWindow(WindowId id) {
            var book = GetBook(id.ProcessId, id.BookName);
            var window = book.Windows[id.WindowIndex];
            Debug.Assert(window != null);
            return window;
        }

        private dynamic GetSheet(SheetId id) {
            var book = GetBook(id.ProcessId, id.BookName);
            object sheet = book.Sheets[id.SheetName];
            Debug.Assert(sheet != null);
            return sheet;
        }
    }
}
