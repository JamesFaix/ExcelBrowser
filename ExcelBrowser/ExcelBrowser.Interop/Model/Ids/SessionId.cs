using System;

namespace ExcelBrowser.Model {

    public class SessionId : IEquatable<SessionId>, IComparable<SessionId> {

        public SessionId(int id) {
            this.Id = id;
        }

        public int Id { get; }

        #region Equality / Comparison

        public bool Equals(SessionId other) => !Equals(other, null)
            && Equals(Id, other.Id);

        public override bool Equals(object obj) => Equals(obj as SessionId);

        public override int GetHashCode() => Id.GetHashCode();

        public int CompareTo(SessionId other) {
            if (Equals(other, null)) return 1;
            return Id.CompareTo(other.Id);
        }

        #endregion

        public override string ToString() => $"Session: {Id}";
    }
}
