using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using xlBook = Microsoft.Office.Interop.Excel.Workbook;
using xlChart = Microsoft.Office.Interop.Excel.Chart;
using xlSheet = Microsoft.Office.Interop.Excel.Worksheet;
using xlWin = Microsoft.Office.Interop.Excel.Window;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a workbook.
    /// </summary>
    public class BookToken : Token<BookId> {

        public BookToken(xlBook book) : base(book?.Id()) {
           // Debug.WriteLine("BookToken.Constructor");

            Sheets = book.Sheets.OfType<dynamic>()
                .Select(s => new SheetToken(s))
                .ToImmutableArray();

            Windows = book.Windows.OfType<xlWin>()
                .Select(w => new WindowToken(w))
                .ToImmutableArray();

            object activeSheet = book.ActiveSheet;
            if (activeSheet is xlSheet) ActiveSheet = new SheetToken(activeSheet as xlSheet);
            else if (activeSheet is xlChart) ActiveSheet = new SheetToken(activeSheet as xlChart);
            else throw new InvalidOperationException("Invalid sheet type.");

        }

        public IEnumerable<SheetToken> Sheets { get; }

        public IEnumerable<WindowToken> Windows { get; }

        public SheetToken ActiveSheet { get; }

        #region Equality

        public bool Equals(BookToken other) => base.Equals(other)
            && Sheets.SequenceEqual(other.Sheets)
            && Windows.SequenceEqual(other.Windows)
            && Equals(ActiveSheet , other.ActiveSheet);

        public override bool Equals(object obj) => Equals(obj as BookToken);

        public bool Matches(BookToken other) => base.Equals(other);

        #endregion
    }
}
