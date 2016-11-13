using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ExcelBrowser.Interop;

namespace ExcelBrowser.Model {

    internal static class TokenFactory {

        public static SessionToken Session(Session session) {
            Requires.NotNull(session, nameof(session));

            var reachableApps = session.Apps
                .Select(App)
                .OrderBy(at => at.Id);

            var unreachableApps = session.UnreachableProcessIds
                .Select(UnreachableApp)
                .OrderBy(at => at.Id);

            var apps = reachableApps
                .Concat(unreachableApps);

            int? topMostProcessId = session.TopMost?.AsProcess()?.Id;
            var activeAppId = topMostProcessId.HasValue
                ? apps.Select(a => a.Id)
                    .SingleOrDefault(id => id.ProcessId == topMostProcessId)
                : null;

            int? primaryProcessId = session.Primary?.AsProcess()?.Id;
            var primaryAppId = primaryProcessId.HasValue
                ? apps.Select(a => a.Id)
                    .SingleOrDefault(id => id.ProcessId == primaryProcessId)
                : null;

            return new SessionToken(
                id: new SessionId(session.SessionId),
                apps: apps,
                activeAppId: activeAppId,
                primaryAppId: primaryAppId);
        }

        public static AppToken App(Application app) {
            Requires.NotNull(app, nameof(app));

            bool isVisible;

            try {
                isVisible = app.Visible
                    && app.AsProcess().IsVisible();
            }
            catch (COMException x)
            when (x.Message.StartsWith("The message filter indicated that the application is busy.")) {
                //This means the application is in a state that does not permit COM automation.
                //Often, this is due to a dialog window or right-click context menu being open.
                Debug.WriteLine($"Busy @ {IdFactory.App(app)}");
                isVisible = false;
            }

            if (isVisible) {
                var activeBook = app.ActiveWorkbook;
                var bookId = activeBook == null ? IdFactory.Book(activeBook) : null;

                var activeWindow = app.ActiveWindow;
                var windowId = activeWindow == null ? IdFactory.Window(activeWindow) : null;

                return new AppToken(
                    id: IdFactory.App(app),
                    isVisible: isVisible,
                    books: app.Workbooks.OfType<Workbook>()
                        .Select(Book),
                    activeBookId: bookId,
                    activeWindowId: windowId);
            }
            else {
                return new AppToken(
                    id: IdFactory.App(app),
                    isVisible: isVisible,
                    books: new BookToken[0],
                    activeBookId: null,
                    activeWindowId: null);
            }
        }

        public static AppToken UnreachableApp(int processId) {
            return new AppToken(
                id: new AppId(processId),
                isVisible: false,
                books: new BookToken[0],
                activeBookId: null,
                activeWindowId: null);
        }

        public static BookToken Book(Workbook book) {
            Requires.NotNull(book, nameof(book));
            return new BookToken(
                id: IdFactory.Book(book),
                isVisible: book.IsVisible(),
                isAddIn: book.IsAddin,
                sheets: book.Sheets.OfType<dynamic>().Select(SheetImpl),
                windows: book.Windows.OfType<Window>().Select(Window),
                activeSheetId: book.ActiveSheet?.Id());
        }

        public static SheetToken Sheet(Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return Sheet(sheet);
        }

        public static SheetToken Sheet(Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return Sheet(chart);
        }

        private static SheetToken SheetImpl(dynamic obj) {
            return new SheetToken(
                id: IdFactory.Sheet(obj),
                isVisible: obj.IsVisible(),
                index: obj.Index);
        }

        public static WindowToken Window(Window win) {
            Requires.NotNull(win, nameof(win));
            return new WindowToken(
                id: IdFactory.Window(win),
                isVisible: win.Visible,
                state: ConvertState(win.WindowState));
        }

        private static WindowState ConvertState(XlWindowState innerState) {
            switch (innerState) {
                case XlWindowState.xlMaximized: return WindowState.Maximized;
                case XlWindowState.xlMinimized: return WindowState.Minimized;
                case XlWindowState.xlNormal: return WindowState.Normal;
                default: throw Requires.ValidEnum((int)innerState);
            }
        }
    }
}
