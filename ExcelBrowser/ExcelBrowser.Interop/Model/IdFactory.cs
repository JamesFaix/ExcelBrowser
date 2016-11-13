using System;
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

        internal static SheetId Sheet(object obj) {
            Requires.NotNull(obj, nameof(obj));

            var sheet = obj as Worksheet;
            if (sheet != null) return Sheet(sheet);

            var chart = obj as Chart;
            if (chart != null) return Sheet(chart);

            throw new NotSupportedException("Invalid sheet type.");
        }

        public static SheetId Sheet(Worksheet sheet) {
            Requires.NotNull(sheet, nameof(sheet));
            return new SheetId(
                processId: sheet.Application.AsProcess().Id,
                bookName: sheet.Workbook().Name,
                sheetName: sheet.Name);
        }

        public static SheetId Sheet(Chart chart) {
            Requires.NotNull(chart, nameof(chart));
            return new SheetId(
                processId: chart.Application.AsProcess().Id,
                bookName: chart.Workbook().Name,
                sheetName: chart.Name);
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
