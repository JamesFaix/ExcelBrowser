using System;
using System.Linq;
using ExcelBrowser.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Model {

    internal static class TokenFactory {
        
        internal static AppToken App(Application app) {
            Requires.NotNull(app, nameof(app));

            return app.IsVisible() 
                ? VisibleApp(app) 
                : InvisibleApp(IdFactory.App(app));
        }

        internal static AppToken VisibleApp(Application app) =>
            new AppToken(
                id: IdFactory.App(app),
                isActive: app.IsActive(),
                isVisible: true,
                version: app.VersionName(),
                books: app.Workbooks.OfType<Workbook>().Select(Book));

        internal static AppToken InvisibleApp(AppId id) =>
            new AppToken(
                id: id,
                isActive: false,
                isVisible: false,
                version: "(Unknown)",
                books: new BookToken[0]);

        internal static BookToken Book(Workbook book) {
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

        internal static SheetToken Sheet(Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return new SheetToken(
                id: IdFactory.Sheet(sheet),
                isActive: sheet.IsActive(),
                isVisible: sheet.IsVisible(),
                index: sheet.Index,
                tabColor: sheet.TabColor());
        }

        internal static SheetToken Sheet(Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return new SheetToken(
                id: IdFactory.Sheet(chart),
                isActive: chart.IsActive(),
                isVisible: chart.IsVisible(),
                index: chart.Index,
                tabColor: chart.TabColor());
        }

        internal static WindowToken Window(Window win) {
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
