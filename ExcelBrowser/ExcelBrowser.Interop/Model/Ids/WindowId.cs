using System;
using System.Runtime.Serialization;

namespace ExcelBrowser.Model {

    [DataContract]
    public class WindowId : IEquatable<WindowId>, IComparable<WindowId> {

        public WindowId(int processId, string bookName, int windowIndex) {
            Requires.NotNull(bookName, nameof(bookName));
            ProcessId = processId;
            BookName = bookName;
            WindowIndex = windowIndex;
        }

        [DataMember(Order =1)]
        public int ProcessId { get; }

        [DataMember(Order = 2)]
        public string BookName { get; }

        [DataMember(Order = 3)]
        public int WindowIndex { get; }

        #region Equality / Comparison

        public bool Equals(WindowId other) => !Equals(other, null)
            && Equals(ProcessId, other.ProcessId)
            && Equals(BookName, other.BookName)
            && Equals(WindowIndex, other.WindowIndex);

        public override bool Equals(object obj) => Equals(obj as WindowId);

        public override int GetHashCode() => WindowIndex.GetHashCode();

        public int CompareTo(WindowId other) {
            if (Equals(other, null)) return 1;

            var x = ProcessId.CompareTo(other.ProcessId);
            if (x != 0) return x;

            x = BookName.CompareTo(other.BookName);
            if (x != 0) return x;

            return WindowIndex.CompareTo(other.WindowIndex);
        }

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
