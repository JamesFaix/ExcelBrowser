using System.Runtime.Serialization;

namespace ExcelBrowser.Model {

    [DataContract]
    public partial class Change {

        internal Change(ChangeType type, object id) {
            Requires.NotNull(id, nameof(id));

            Type = type;
            Id = id;
        }

        [DataMember(Order = 1)]
        public ChangeType Type { get; }

        [DataMember(Order = 2)]
        public object Id { get; }

        public override string ToString() => Serializer.Serialize(this);
    }

    public class Change<TId> : Change {
        internal Change(ChangeType type, TId id)
            : base(type, id) { }

        public new TId Id => (TId)base.Id;

        public override string ToString() => Serializer.Serialize(this);
    }

    [DataContract]
    public class Change<TId, TValue> : Change<TId> {
        internal Change(ChangeType type, TId id, TValue value)
            : base(type, id) {
            Value = value;
        }

        [DataMember(Order = 3)]
        public TValue Value { get; }

        public override string ToString() => Serializer.Serialize(this);
    }
}
