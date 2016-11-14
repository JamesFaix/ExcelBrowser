using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Interop {

    public static class WorkbookExtensionMethods {
        
        public static bool IsVisible(this Workbook book) {
            Requires.NotNull(book, nameof(book));
            return book.Windows.OfType<Window>()
                .Where(wn => wn.Visible)
                .Any();
        }

        public static bool IsActive(this Workbook book) {
            Requires.NotNull(book, nameof(book));
            return Equals(book, book.Application.ActiveWorkbook);
        }
    }
}
