using System;
using System.Runtime.Serialization;

namespace ExcelBrowser.Model {

    [DataContract]
    public class BookId : IEquatable<BookId>, IComparable<BookId> {

        public BookId(int processId, string bookName) {
            Requires.NotNull(bookName, nameof(bookName));
            ProcessId = processId;
            BookName = bookName;
        }

        [DataMember(Order = 1)]
        public int ProcessId { get; }

        [DataMember(Order = 2)]
        public string BookName { get; }

        #region Equality / Comparison

        public bool Equals(BookId other) => !Equals(other, null)
            && Equals(ProcessId, other?.ProcessId)
            && Equals(BookName, other?.BookName);

        public override bool Equals(object obj) => Equals(obj as BookId);

        public override int GetHashCode() => BookName.GetHashCode();

        public int CompareTo(BookId other) {
            if (Equals(other, null)) return 1;

            var x = ProcessId.CompareTo(other.ProcessId);
            if (x != 0) return x;

            return BookName.CompareTo(other.BookName);
        }

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
