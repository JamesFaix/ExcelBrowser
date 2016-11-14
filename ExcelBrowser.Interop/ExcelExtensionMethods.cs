using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Windows.Media;
using System.Runtime.InteropServices;

namespace ExcelBrowser.Interop {

    public static class ExcelExtensionMethods {

        #region Get Workbook

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

        #endregion

        #region IsVisible

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

        public static bool IsVisible(this Application app) {
            Requires.NotNull(app, nameof(app));
            
            try {
                return app.Visible
                    && app.AsProcess().IsVisible();
            }
            catch (COMException x)
            when (x.Message.StartsWith("The message filter indicated that the application is busy.")
                || x.Message.StartsWith("Call was rejected by callee.")) {
                //This means the application is in a state that does not permit COM automation.
                //Often, this is due to a dialog window or right-click context menu being open.
                return false;
            }
        }
        #endregion

        #region IsActive

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

        public static bool IsActive(this Application app) {
            Requires.NotNull(app, nameof(app));
            return Equals(app, app.Session().TopMost);
        }

        #endregion

        public static Session Session(this Application app) {
            Requires.NotNull(app, nameof(app));
            return new Session(app.AsProcess().SessionId);
        }

        public static Color? TabColor(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));

            var result = sheet.Tab.Color;
            return Equals(result, false)
                ? null
                : ColorTranslator.FromOle(result);
        }

        public static Color TabColor(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));

            var result = chart.Tab.Color;
            return Equals(result, false)
                ? null
                : ColorTranslator.FromOle(result);
        }
    }
}
