using System.Windows.Media;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Interop {

    public static class SheetExtensionMethods {

        public static Workbook Workbook(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return sheet.Parent as Workbook;
        }

        public static Workbook Workbook(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return chart.Parent as Workbook;
        }
        public static bool IsVisible(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return sheet.Visible == XlSheetVisibility.xlSheetVisible;
        }

        public static bool IsVisible(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return chart.Visible == XlSheetVisibility.xlSheetVisible;
        }
        
        public static bool IsActive(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return Equals(sheet, sheet.Workbook().ActiveSheet);
        }

        public static bool IsActive(this Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return Equals(chart, chart.Workbook().ActiveSheet);
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
