using System;

namespace ExcelBrowser.Model {

    public class SheetId : IEquatable<SheetId>, IComparable<SheetId> {

        public SheetId(int processId, string bookName, string sheetName) {
            Requires.NotNull(bookName, nameof(bookName));
            Requires.NotNull(sheetName, nameof(sheetName));
            ProcessId = processId;
            BookName = bookName;
            SheetName = sheetName;
        }

        public int ProcessId { get; }
        public string BookName { get; }
        public string SheetName { get; }

        #region Equality / Comparison

        public bool Equals(SheetId other) => !Equals(other, null)
            && Equals(ProcessId, other.ProcessId)
            && Equals(BookName, other.BookName)
            && Equals(SheetName, other.SheetName);

        public override bool Equals(object obj) => Equals(obj as SheetId);

        public override int GetHashCode() => SheetName.GetHashCode();

        public int CompareTo(SheetId other) {
            if (Equals(other, null)) return 1;

            var x = ProcessId.CompareTo(other.ProcessId);
            if (x != 0) return x;

            x = BookName.CompareTo(other.BookName);
            if (x != 0) return x;

            return SheetName.CompareTo(other.SheetName);
        }

        #endregion

        public override string ToString() => $"{{Process: {ProcessId}, Book: {BookName}, Sheet: {SheetName}}}";
    }
}
