using System;
using System.Runtime.Serialization;
using ExcelBrowser.Interop;
using xlChart = Microsoft.Office.Interop.Excel.Chart;
using xlSheet = Microsoft.Office.Interop.Excel.Worksheet;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a sheet.
    /// </summary>
    [DataContract]
    public class SheetToken : Token<SheetId>, IEquatable<SheetToken> {

        public SheetToken(xlSheet sheet) : base(sheet?.Id()) {
            IsVisible = sheet.IsVisible();
            //Debug.WriteLine("SheetToken.Constructor");
        }

        public SheetToken(xlChart chart) : base(chart?.Id()) {
            IsVisible = chart.IsVisible();
        }

        [DataMember(Order = 1)]
        public bool IsVisible { get; }

        #region Equality

        public bool Matches(SheetToken other) => base.Equals(other);

        public bool Equals(SheetToken other) => base.Equals(other)
            && IsVisible == other.IsVisible;

        public override bool Equals(object obj) => Equals(obj as SheetToken);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}