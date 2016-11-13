using System;
using System.Runtime.Serialization;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of an Excel window.
    /// </summary>
    [DataContract]
    public class WindowToken : Token<WindowId>, IEquatable<WindowToken> {

        public WindowToken(WindowId id, bool isActive, bool isVisible, WindowState state) 
            : base(id) {
            IsActive = isActive;
            IsVisible = isVisible;
            State = state;
        }
        
        [DataMember(Order = 2)]
        public bool IsActive { get; }

        [DataMember(Order = 3)]
        public bool IsVisible { get; }

        [DataMember(Order = 4)]
        public WindowState State { get; }

        #region Equality

        public bool Matches(WindowToken other) => base.Equals(other);

        public bool Equals(WindowToken other) => base.Equals(other)
            && IsActive == other.IsActive
            && IsVisible == other.IsVisible
            && State == other.State;

        public override bool Equals(object obj) => Equals(obj as WindowToken);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
