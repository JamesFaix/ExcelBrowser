namespace ExcelBrowser.Model {

    partial class Change {

        private const string START = "Start";
        private const string ADD = "Added";
        private const string REMOVE = "Removed";
        private const string ACTIVATE = "Activated";
        private const string VISIBLE = "Visible";
        private const string STATE = "State";
        private const string MOVE = "Move";

        public static Change Added<TId>(TId id) =>
            new Change<TId>(ADD, id);

        public static Change Removed<TId>(TId id) =>
            new Change<TId>(REMOVE, id);

        public static Change Activated<TId>(TId id) =>
            new Change<TId>(ACTIVATE, id);

        public static Change SetVisibility<TId>(TId id, bool value) =>
            new Change<TId, bool>(VISIBLE, id, value);

        public static Change SessionStart(SessionId id) =>
            new Change<SessionId>(START, id);

        public static Change WindowSetState(WindowId id, WindowState value) =>
            new Change<WindowId, WindowState>(STATE, id, value);

        public static Change SheetMove(SheetId id, int index) =>
            new Change<SheetId, int>(MOVE, id, index);
    }
}
