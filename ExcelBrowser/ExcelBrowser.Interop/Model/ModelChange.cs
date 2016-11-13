using System.Runtime.Serialization;

namespace ExcelBrowser.Model {

    [DataContract]
    public partial class ModelChange {

        internal ModelChange(string type, object id) {
            Requires.NotNull(type, nameof(type));
            Requires.NotNull(id, nameof(id));

            Type = type;
            Id = id;
        }

        [DataMember(Order = 1)]
        public string Type { get; }

        [DataMember(Order = 2)]
        public object Id { get; }

        public override string ToString() => Serializer.Serialize(this);
    }

    public class ModelChange<TId> : ModelChange {
        internal ModelChange(string type, TId id)
            : base(type, id) { }

        public new TId Id => (TId)base.Id;

        public override string ToString() => Serializer.Serialize(this);
    }

    [DataContract]
    public class ModelChange<TId, TValue> : ModelChange<TId> {
        internal ModelChange(string type, TId id, TValue value)
            : base(type, id) {
            Value = value;
        }

        [DataMember(Order = 3)]
        public TValue Value { get; }

        public override string ToString() => Serializer.Serialize(this);
    }
}
