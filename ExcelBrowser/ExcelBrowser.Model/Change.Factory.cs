namespace ExcelBrowser.Model {

    partial class Change {
        
        public static Change Added<TId>(TId id) =>
            new Change<TId>("Add", id);

        public static Change Removed<TId>(TId id) =>
            new Change<TId>("Remove", id);

        public static Change SetActive<TId>(TId id, bool value) =>
            new Change<TId, bool>("SetActive", id, value);

        public static Change SetVisibility<TId>(TId id, bool value) =>
            new Change<TId, bool>("SetVisible", id, value);

        public static Change SessionStart(SessionId id) =>
            new Change<SessionId>("Start", id);

        public static Change WindowSetState(WindowId id, WindowState value) =>
            new Change<WindowId, WindowState>("SetState", id, value);

        public static Change SheetMove(SheetId id, int index) =>
            new Change<SheetId, int>("SetIndex", id, index);
    }
}
