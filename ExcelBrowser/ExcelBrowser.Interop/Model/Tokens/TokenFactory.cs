using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelBrowser.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Model {

    public class TokenFactory {

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
                Debug.WriteLine($"Busy @ {app.Id()}");
                isVisible = false;
            }

            if (isVisible) {
                return new AppToken(
                    id: app.Id(),
                    isVisible: isVisible,
                    books: app.Workbooks.OfType<Workbook>()
                        .Select(Book),
                    activeBookId: app.ActiveWorkbook?.Id(),
                    activeWindowId: app.ActiveWindow?.Id());
            }
            else {
                return new AppToken(
                    id: app.Id(),
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
                id: book.Id(),
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
                id: obj.Id(),
                isVisible: obj.IsVisible(),
                index: obj.Index);
        }

        public static WindowToken Window(Window win) {
            Requires.NotNull(win, nameof(win));
            return new WindowToken(
                id: win.Id(),
                isVisible: win.Visible,
                state: win.WindowState.Outer());
        }

    }
}
