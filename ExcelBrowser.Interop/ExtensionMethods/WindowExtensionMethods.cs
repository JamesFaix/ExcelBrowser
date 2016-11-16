using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Interop {

    public static class WindowExtensionMethods {

        public static Workbook Workbook(this Window window) {
            Requires.NotNull(window, nameof(window));
            return window.Parent as Workbook;
        }

        public static bool IsActive(this Window window) {
            Requires.NotNull(window, nameof(window));

            var other = window.Application.ActiveWindow;
            if (other == null) return false;

            return window.WindowNumber == other.WindowNumber
                && window.Workbook().Name == other.Workbook().Name;
        }
    }
}
