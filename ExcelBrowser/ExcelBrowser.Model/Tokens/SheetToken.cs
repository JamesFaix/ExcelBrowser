using System;
using System.Runtime.Serialization;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a sheet.
    /// </summary>
    [DataContract]
    public class SheetToken : Token<SheetId>, IEquatable<SheetToken> {

        public SheetToken(SheetId id, bool isVisible, int index) 
            : base(id) {
            IsVisible = isVisible;
            Index = index;
        }
        
        [DataMember(Order = 2)]
        public bool IsVisible { get; }

        [DataMember(Order = 3)]
        public int Index { get; }

        #region Equality

        public bool Matches(SheetToken other) => base.Equals(other);

        public bool Equals(SheetToken other) => base.Equals(other)
            && IsVisible == other.IsVisible
            && Index == other.Index;

        public override bool Equals(object obj) => Equals(obj as SheetToken);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}