using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.Serialization;

#pragma warning disable CS0659 
//Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of a workbook.
    /// </summary>
    [DataContract]
    public class BookToken : Token<BookId> {

        public BookToken(BookId id, bool isVisible, bool isAddIn, 
            IEnumerable<SheetToken> sheets, IEnumerable<WindowToken> windows,
            SheetId activeSheetId) 
            : base(id) {
            Requires.NotNull(sheets, nameof(sheets));
            Requires.NotNull(windows, nameof(windows));

            IsVisible = isVisible;
            IsAddIn = isAddIn;
            Sheets = sheets.ToImmutableArray();
            Windows = windows.ToImmutableArray();

            if (activeSheetId != null) {
                try {
                    ActiveSheet = Sheets.Single(s => Equals(s.Id, activeSheetId));
                }
                catch (InvalidOperationException x)
                when (x.Message.StartsWith("Sequence contains no elements")) {
                    throw new InvalidOperationException("ActiveSheet ID not found in sheets collection.", x);
                }
            }
        }
        
        [DataMember(Order = 2)]
        public bool IsVisible { get; }

        [DataMember(Order = 3)]
        public bool IsAddIn { get; }

        [DataMember(Order = 4)]
        public IEnumerable<SheetToken> Sheets { get; }

        [DataMember(Order = 5)]
        public SheetToken ActiveSheet { get; }

        [DataMember(Order = 6)]
        public IEnumerable<WindowToken> Windows { get; }

        #region Equality

        public bool Equals(BookToken other) => base.Equals(other)
            && Sheets.SequenceEqual(other.Sheets)
            && Windows.SequenceEqual(other.Windows)
            && Equals(ActiveSheet, other.ActiveSheet)
            && IsVisible == other.IsVisible
            && IsAddIn == other.IsAddIn;

        public override bool Equals(object obj) => Equals(obj as BookToken);

        public bool Matches(BookToken other) => base.Equals(other);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
