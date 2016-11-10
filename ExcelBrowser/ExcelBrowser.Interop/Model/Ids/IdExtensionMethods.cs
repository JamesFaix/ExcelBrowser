using ExcelBrowser.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Model {

    public static class IdExtensionMethods {

        public static AppId Id(this Application app) {
            Requires.NotNull(app, nameof(app));
            return new AppId(app.AsProcess().Id);
        }

        public static BookId Id(this Workbook book) {
            Requires.NotNull(book, nameof(book));
            return new BookId(book.Application.AsProcess().Id, book.Name);
        }

        public static SheetId Id(this Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return new SheetId(sheet.Application.AsProcess().Id, sheet.Workbook().Name, sheet.Name);
        }

        public static SheetId Id(this Chart sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return new SheetId(sheet.Application.AsProcess().Id, sheet.Workbook().Name, sheet.Name);
        }
        
        public static WindowId Id(this Window window) {
            Requires.NotNull(window, nameof(window));
            return new WindowId(window.Application.AsProcess().Id, window.Workbook().Name, window.WindowNumber);
        }
    }
}
