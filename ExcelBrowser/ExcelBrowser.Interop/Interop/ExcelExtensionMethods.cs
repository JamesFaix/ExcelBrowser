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

        //public static int Index(this Workbook book) {
        //    Requires.NotNull(book, nameof(book));
        //    return book.Application.Workbooks.OfType<Workbook>().IndexOf(book) + 1;
        //}
        

    }
}
