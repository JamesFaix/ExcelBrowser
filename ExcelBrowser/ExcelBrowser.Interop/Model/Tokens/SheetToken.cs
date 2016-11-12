using System;
using xlChart = Microsoft.Office.Interop.Excel.Chart;
using xlSheet = Microsoft.Office.Interop.Excel.Worksheet;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a sheet.
    /// </summary>
    public class SheetToken : Token<SheetId>, IEquatable<SheetToken> {

        public SheetToken(xlSheet sheet) : base(sheet?.Id()) {
            //Debug.WriteLine("SheetToken.Constructor");
        }

        public SheetToken(xlChart chart) : base(chart?.Id()) {
        }

        public bool Equals(SheetToken other) => base.Equals(other);

        public override bool Equals(object obj) => Equals(obj as SheetToken);
    }
}