using ExcelBrowser.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelBrowser.Model {

    internal static class IdFactory {

        public static AppId App(Application app) {
            Requires.NotNull(app, nameof(app));
            return new AppId(
                processId: app.AsProcess().Id);
        }

        public static BookId Book(Workbook book) {
            Requires.NotNull(book, nameof(book));
            return new BookId(
                processId: book.Application.AsProcess().Id, 
                bookName: book.Name);
        }

        public static SheetId Sheet(Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return SheetImpl(sheet);
        }

        public static SheetId Sheet(Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return SheetImpl(chart);
        }

        internal static SheetId SheetImpl(dynamic obj) {
            return new SheetId(
                processId: obj.Application.AsProcess().Id,
                bookName: obj.Workbook().Name,
                sheetName: obj.Name);
        }
        
        public static WindowId Window(Window window) {
            Requires.NotNull(window, nameof(window));
            return new WindowId(
                processId: window.Application.AsProcess().Id, 
                bookName: window.Workbook().Name, 
                windowIndex: window.WindowNumber);
        }
    }
}
