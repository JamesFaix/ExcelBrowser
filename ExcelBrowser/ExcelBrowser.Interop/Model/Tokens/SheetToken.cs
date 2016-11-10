using xlChart = Microsoft.Office.Interop.Excel.Chart;
using xlSheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Diagnostics;
using System;

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a sheet.
    /// </summary>
    public class SheetToken : Token<SheetId>, IEquatable<SheetToken> {

        public SheetToken(xlSheet sheet) : base(sheet?.Id()) {
        //    Debug.WriteLine("SheetToken.Constructor");
        }

        public SheetToken(xlChart chart) : base(chart?.Id()) {
        }

        public bool Equals(SheetToken other) => base.Equals(other);

        public override bool Equals(object obj) => Equals(obj as SheetToken);
    }
}
