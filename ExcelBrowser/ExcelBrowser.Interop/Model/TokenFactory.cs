using System;
using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Model {

    internal class TokenFactory {

        public SessionToken Session(Session session) {
            Requires.NotNull(session, nameof(session));

            IEnumerable<AppToken> reachableApps = session.Apps
                .Select(App)
                .OrderBy(at => at.Id);

            var unreachableApps = session.UnreachableProcessIds
                .Select(id => InvisibleApp(new AppId(id)))
                .OrderBy(a => a.Id);

            var apps = ReplaceInvisibleApps(
                reachableApps.Concat(unreachableApps));

            var primary = session.Primary;
            var primaryId = (primary != null) ? IdFactory.App(primary) : null;

            var result = new SessionToken(
                id: new SessionId(session.SessionId),
                apps: apps,
                primaryAppId: primaryId);

            this.previousSession = result;
            return result;
        }

        #region Freezing busy apps

        private SessionToken previousSession;

        //Replaces any busy apps that were previously cached with their previous version
        private IEnumerable<AppToken> ReplaceInvisibleApps(IEnumerable<AppToken> apps) {
            if (this.previousSession == null) {
                //Don't do anything special before first snapshot is saved
                foreach (var a in apps) yield return a;
            }
            else {
                var previousApps = this.previousSession.Apps.ToArray();

                foreach (var a in apps) {
                    if(a.IsVisible) { //If visible, return input
                        yield return a;
                    }
                    else { //Return previous match, but default to input
                        var prev = previousApps.SingleOrDefault(p => Equals(p.Id, a.Id));
                        if (prev != null) { //Mark and return previous match
                            prev = prev.ShallowCopy;
                            prev.IsVisible = false;
                            yield return prev;
                        }
                        else { //If no previous match, return input
                            yield return a;
                        }
                    }                    
                }
            }
        }

        #endregion

        private static AppToken App(Application app) {
            Requires.NotNull(app, nameof(app));

            return app.IsVisible() 
                ? VisibleApp(app) 
                : InvisibleApp(IdFactory.App(app));
        }

        private static AppToken VisibleApp(Application app) =>
            new AppToken(
                id: IdFactory.App(app),
                isActive: app.IsActive(),
                isVisible: true,
                books: app.Workbooks.OfType<Workbook>().Select(Book));
        
        private static AppToken InvisibleApp(AppId id) =>
            new AppToken(
                id: id,
                isActive: false,
                isVisible: false,
                books: new BookToken[0]);

        private static BookToken Book(Workbook book) {
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

        private static SheetToken Sheet(Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return new SheetToken(
                id: IdFactory.Sheet(sheet),
                isActive: sheet.IsActive(),
                isVisible: sheet.IsVisible(),
                index: sheet.Index,
                tabColor: sheet.TabColor());
        }

        private static SheetToken Sheet(Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return new SheetToken(
                id: IdFactory.Sheet(chart),
                isActive: chart.IsActive(),
                isVisible: chart.IsVisible(),
                index: chart.Index,
                tabColor: chart.TabColor());
        }

        private static WindowToken Window(Window win) {
            Requires.NotNull(win, nameof(win));

            var activeSheet = win.ActiveSheet;
            var activeSheetId = activeSheet == null ? IdFactory.Sheet(activeSheet) : null;

            return new WindowToken(
                id: IdFactory.Window(win),
                isActive: win.IsActive(),
                isVisible: win.Visible,
                state: ConvertState(win.WindowState),
                activeSheetId: activeSheetId);
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
