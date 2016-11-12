using System;

namespace ExcelBrowser.Model {

    public class AppId : IEquatable<AppId>, IComparable<AppId> {

        public AppId(int processId) {
            this.ProcessId = processId;
        }

        public int ProcessId { get; }

        #region Equality / Comparison

        public bool Equals(AppId other) => !Equals(other, null)
            && ProcessId == other.ProcessId;

        public override bool Equals(object obj) => Equals(obj as AppId);

        public override int GetHashCode() => ProcessId.GetHashCode();

        public int CompareTo(AppId other) {
            if (other == null) return 1;
            return ProcessId.CompareTo(other.ProcessId);
        }

        #endregion

        public override string ToString() => $"{{Process: {ProcessId}}}";
    }
}
