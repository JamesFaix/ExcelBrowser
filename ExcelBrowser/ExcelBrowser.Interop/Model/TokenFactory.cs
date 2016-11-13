using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ExcelBrowser.Interop;
using System;

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

            return new AppToken(
                id: IdFactory.App(app),
                isActive: app.IsActive(),
                isVisible: isVisible,
                books: isVisible
                    ? app.Workbooks.OfType<Workbook>().Select(Book)
                    : new BookToken[0]);
        }

        public static AppToken UnreachableApp(int processId) {
            return new AppToken(
                id: new AppId(processId),
                isActive: false,
                isVisible: false,
                books: new BookToken[0]);
        }

        public static BookToken Book(Workbook book) {
            Requires.NotNull(book, nameof(book));
            return new BookToken(
                id: IdFactory.Book(book),
                isActive: book.IsActive(),
                isVisible: book.IsVisible(),
                isAddIn: book.IsAddin,
                sheets: book.Sheets.OfType<object>().Select(Sheet),
                windows: book.Windows.OfType<Window>().Select(Window));
        }

        private static SheetToken Sheet(object obj) {
            var sheet = obj as Worksheet;
            if (obj != null) return Sheet(sheet);

            var chart = obj as Chart;
            if (obj != null) return Sheet(chart);

            throw new NotSupportedException("Invalid sheet type.");
        }

        public static SheetToken Sheet(Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return new SheetToken(
                id: IdFactory.Sheet(sheet),
                isActive: sheet.IsActive(),
                isVisible: sheet.IsVisible(),
                index: sheet.Index);
        }

        public static SheetToken Sheet(Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return new SheetToken(
                id: IdFactory.Sheet(chart),
                isActive: chart.IsActive(),
                isVisible: chart.IsVisible(),
                index: chart.Index);
        }
        
        public static WindowToken Window(Window win) {
            Requires.NotNull(win, nameof(win));
            return new WindowToken(
                id: IdFactory.Window(win),
                isActive: win.IsActive(),
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
