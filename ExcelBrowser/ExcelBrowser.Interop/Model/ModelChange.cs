using System.Diagnostics;

namespace ExcelBrowser.Model {

    public abstract class ModelChange {

        protected ModelChange(object id, ModelChangeType type) {
            Requires.NotNull(id, nameof(id));
            Id = id;
            Type = type;
            Debug.WriteLine("ModelChange.Constructor: " + this.ToString());
        }

        public object Id { get; }
        public ModelChangeType Type { get; }

        public override string ToString() => $"{Type} @ {Id}";

        public static ModelChange Create<TId>(TId id, ModelChangeType type) =>
            new ModelChange<TId>(id, type);
    }

    public class ModelChange<TId> : ModelChange {

        public ModelChange(TId id, ModelChangeType type)
            : base(id, type) { }

        public new TId Id => (TId)base.Id;
    }    
}
