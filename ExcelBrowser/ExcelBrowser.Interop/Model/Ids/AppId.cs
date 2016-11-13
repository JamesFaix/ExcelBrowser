using System;
using System.Runtime.Serialization;

namespace ExcelBrowser.Model {

    [DataContract]
    public class AppId : IEquatable<AppId>, IComparable<AppId> {

        public AppId(int processId) {
            this.ProcessId = processId;
        }

        [DataMember]
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

        public override string ToString() => Serializer.Serialize(this);
    }
}
