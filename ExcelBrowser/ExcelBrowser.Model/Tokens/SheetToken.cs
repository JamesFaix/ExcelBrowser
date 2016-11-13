using System;
using System.Runtime.Serialization;
using System.Drawing;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a sheet.
    /// </summary>
    [DataContract]
    public class SheetToken : Token<SheetId>, IEquatable<SheetToken> {

        public SheetToken(SheetId id, bool isActive, bool isVisible, int index, Color tabColor) 
            : base(id) {

            IsActive = isActive;
            IsVisible = isVisible;
            Index = index;
            TabColor = tabColor;
        }
        
        [DataMember(Order = 2)]
        public bool IsActive { get; }

        [DataMember(Order = 3)]
        public bool IsVisible { get; }

        [DataMember(Order = 4)]
        public int Index { get; }

        [DataMember(Order = 5)]
        public Color TabColor { get; } 

        #region Equality

        public bool Matches(SheetToken other) => base.Equals(other);

        public bool Equals(SheetToken other) => base.Equals(other)
            && IsActive == other.IsActive
            && IsVisible == other.IsVisible
            && Index == other.Index
            && TabColor == other.TabColor;

        public override bool Equals(object obj) => Equals(obj as SheetToken);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}