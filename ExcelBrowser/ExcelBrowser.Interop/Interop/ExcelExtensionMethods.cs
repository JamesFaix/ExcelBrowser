using Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ExcelBrowser.Interop {

    public static class ExcelExtensionMethods {

        public static Workbook Workbook(this Window window) {
            Requires.NotNull(window, nameof(window));
            return window.Parent as Workbook;
        }

        public static Workbook Workbook(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return sheet.Parent as Workbook;
        }

        public static Workbook Workbook(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return chart.Parent as Workbook;
        }
        
        public static bool IsVisible(this Workbook book) {
            Requires.NotNull(book, nameof(book));
            return book.Windows.OfType<Window>()
                .Where(wn => wn.Visible)
                .Any();
        }

        public static bool IsVisible(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return sheet.Visible == XlSheetVisibility.xlSheetVisible;
        }

        public static bool IsVisible(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return chart.Visible == XlSheetVisibility.xlSheetVisible;
        }

        public static bool IsActive(this Window window) {
            Requires.NotNull(window, nameof(window));
            return Equals(window, window.Application.ActiveWindow);
        }

        public static bool IsActive(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return Equals(sheet, sheet.Workbook().ActiveSheet);
        }

        public static bool IsActive(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return Equals(chart, chart.Workbook().ActiveSheet);
        }

        public static bool IsActive(this Workbook book) {
            Requires.NotNull(book, nameof(book));
            return Equals(book, book.Application.ActiveWorkbook);
        }

        public static Session Session(this Application app) {
            Requires.NotNull(app, nameof(app));
            return new Session(app.AsProcess().SessionId);
        }

        public static bool IsActive(this Application app) {
            Requires.NotNull(app, nameof(app));
            return Equals(app, app.Session().TopMost);
        }
    }
}
