namespace ExcelBrowser.Model {

    partial class ModelChange {

        private const string START = "Start";
        private const string ADD = "Added";
        private const string REMOVE = "Removed";
        private const string ACTIVATE = "Activated";
        private const string VISIBLE = "Visible";
        private const string STATE = "State";
        private const string MOVE = "Move";

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

        public static ModelChange WindowSetState(WindowId id, WindowState value) =>
            new ModelChange<WindowId, WindowState>(STATE, id, value);

        public static ModelChange SheetMove(SheetId id, int index) =>
            new ModelChange<SheetId, int>(MOVE, id, index);
    }
}
