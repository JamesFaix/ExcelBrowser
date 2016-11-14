using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.Serialization;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of an Excel application instance.
    /// </summary>
    [DataContract]
    public class AppToken : Token<AppId> {

        public AppToken(AppId id, bool isActive, bool isVisible,
            IEnumerable<BookToken> books) 
            : base(id) {
            Requires.NotNull(books, nameof(books));

            IsActive = isActive;
            IsVisible = isVisible;
            Books = books.ToImmutableArray();            
        }

        public AppToken ShallowCopy => MemberwiseClone() as AppToken;

        [DataMember(Order = 2)]
        public bool IsActive { get; }

        [DataMember(Order = 3)]
        public bool IsVisible { get; set; } //Must be settable for caching

        [DataMember(Order = 4)]
        public IEnumerable<BookToken> Books { get; private set; }

        #region Equality

        public bool Equals(AppToken other) => base.Equals(other)
            && IsActive == other.IsActive
            && IsVisible == other.IsVisible
            && Books.SequenceEqual(other.Books);

        public override bool Equals(object obj) => Equals(obj as AppToken);

        public bool Matches(AppToken other) => base.Equals(other);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
