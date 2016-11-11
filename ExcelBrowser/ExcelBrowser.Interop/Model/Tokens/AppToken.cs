using System;
using System.Collections.Immutable;
using System.Linq;
using ExcelBrowser.Interop;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using xlBook = Microsoft.Office.Interop.Excel.Workbook;
using xlWin = Microsoft.Office.Interop.Excel.Window;
using System.Diagnostics;
using System.Collections.Generic;

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of an Excel application instance.
    /// </summary>
    public class AppToken : Token<AppId> {

        public AppToken(xlApp app) : this(app?.Id()) {
            //    Debug.WriteLine("AppToken.Constructor");

            IsReachable = true;

            Books = app.Workbooks.OfType<xlBook>()
                .Select(wb => new BookToken(wb))
                .ToImmutableArray();

            xlBook activeBook = app.ActiveWorkbook;
            if (activeBook != null) {
                var id = activeBook.Id();
                ActiveBook = Books.Single(b => Equals(b.Id, id));
            }

            xlWin activeWindow = app.ActiveWindow;
            if (activeWindow != null) {
                var id = activeWindow.Id();
                ActiveWindow = Books.Single(b => Equals(b.Id.BookName, id.BookName))
                    .Windows.Single(w => Equals(w.Id, id));
            }
        }

        private AppToken(AppId id) : base(id) { }

        public static AppToken Unreachable(int processId) {
            return new AppToken(new AppId(processId)) {
                Books = new BookToken[0].ToImmutableArray()
            };   
        }

        public bool IsReachable { get; }

        public IEnumerable<BookToken> Books { get; private set; }

        public BookToken ActiveBook { get; }

        public WindowToken ActiveWindow { get; }

        #region Equality

        public bool Equals(AppToken other) => base.Equals(other)
            && IsReachable.Equals(other.IsReachable)
            && Books.SequenceEqual(other.Books)
            && Equals(ActiveBook, other.ActiveBook)
            && Equals(ActiveWindow, other.ActiveWindow);

        public override bool Equals(object obj) => Equals(obj as AppToken);

        public bool Matches(AppToken other) => base.Equals(other);        

        #endregion
    }
}
