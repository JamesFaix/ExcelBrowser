namespace ExcelBrowser.Model {

    public class ModelChange {

        internal ModelChange(string type, object id) {
            Requires.NotNull(type, nameof(type));
            Requires.NotNull(id, nameof(id));

            Type = type;
            Id = id;
        }

        public string Type { get; }
        public object Id { get; }

        public override string ToString() => $"{Type} @ {Id}";

        #region Factory methods

        private const string START = "Start";
        private const string ADD = "Added";
        private const string REMOVE = "Removed";
        private const string ACTIVATE = "Activated";
        private const string VISIBLE = "SetVisibility";
        private const string REACHABLE = "SetReachability";
        private const string STATE = "SetState";

        public static ModelChange Added<TId>(TId id) =>
            new ModelChange<TId>(ADD, id);

        public static ModelChange Removed<TId>(TId id) => 
            new ModelChange<TId>(REMOVE, id);

        public static ModelChange Activated<TId>(TId id) => 
            new ModelChange<TId>(ACTIVATE, id);

        public static ModelChange SetVisibility<TId>(TId id, bool value) => 
            new ModelChange<TId, bool>(VISIBLE, id, value);

        public static ModelChange SessionStart(SessionId id) => 
            new ModelChange<SessionId>(START, id);

        public static ModelChange AppSetReachablity(AppId id, bool value) => 
            new ModelChange<AppId, bool>(REACHABLE, id, value);

        public static ModelChange WindowSetState(WindowId id, WindowState value) => 
            new ModelChange<WindowId, WindowState>(STATE, id, value);

        #endregion
    }

    public class ModelChange<TId> : ModelChange {
        internal ModelChange(string type, TId id)
            : base(type, id) { }

        public new TId Id => (TId)base.Id;
    }

    public class ModelChange<TId, TValue> : ModelChange<TId> {
        internal ModelChange(string type, TId id, TValue value)
            : base(type, id) {
            Value = value;
        }

        public TValue Value { get; }

        public override string ToString() => $"{Type} @ {Id} | {Value}";
    }
}
