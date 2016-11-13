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

        public AppToken(AppId id, bool isVisible, 
            IEnumerable<BookToken> books, BookId activeBookId, WindowId activeWindowId) 
            : base(id) {
            Requires.NotNull(books, nameof(books));

            IsVisible = isVisible;
            Books = books.ToImmutableArray();

            if (activeBookId != null) {
                try {
                    ActiveBook = Books.Single(b => Equals(b.Id, activeBookId));
                }
                catch (InvalidOperationException x)
                when (x.Message.StartsWith("Sequence contains no elements")) {
                    throw new InvalidOperationException("ActiveBook ID not found in books collection.", x);
                }
            }

            if (activeWindowId != null) {
                try {
                    ActiveWindow = Books
                        .SelectMany(b => b.Windows)
                        .Single(w => Equals(w.Id, activeWindowId));
                }
                catch (InvalidOperationException x)
                when (x.Message.StartsWith("Sequence contains no elements")) {
                    throw new InvalidOperationException("ActiveWindow ID not found in windows collection.", x);
                }
            }
        }
        
        [DataMember(Order = 2)]
        public bool IsVisible { get; }

        [DataMember(Order = 3)]
        public IEnumerable<BookToken> Books { get; private set; }

        [DataMember(Order = 4)]
        public BookToken ActiveBook { get; }

        [DataMember(Order = 5)]
        public WindowToken ActiveWindow { get; }

        #region Equality

        public bool Equals(AppToken other) => base.Equals(other)
            && IsVisible == other.IsVisible
            && Books.SequenceEqual(other.Books)
            && Equals(ActiveBook, other.ActiveBook)
            && Equals(ActiveWindow, other.ActiveWindow);

        public override bool Equals(object obj) => Equals(obj as AppToken);

        public bool Matches(AppToken other) => base.Equals(other);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
