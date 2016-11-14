using System.Drawing;

namespace ExcelBrowser.Model {

    partial class Change {
        
        public static Change Added<TParentId, TChildId>(TParentId parentId, TChildId childId) =>
            new Change<TParentId, TChildId>(ChangeType.Add, parentId, childId);

        public static Change Removed<TParentId, TChildId>(TParentId parentId, TChildId childId) =>
            new Change<TParentId, TChildId>(ChangeType.Remove, parentId, childId);

        public static Change SetActive<TId>(TId id, bool value) =>
            new Change<TId, bool>(ChangeType.SetActive, id, value);

        public static Change SetVisibility<TId>(TId id, bool value) =>
            new Change<TId, bool>(ChangeType.SetVisible, id, value);

        public static Change SessionStart(SessionId id) =>
            new Change<SessionId>(ChangeType.SessionStart, id);

        public static Change WindowSetState(WindowId id, WindowState value) =>
            new Change<WindowId, WindowState>(ChangeType.SetWindowState, id, value);

        public static Change SheetSetIndex(SheetId id, int index) =>
            new Change<SheetId, int>(ChangeType.SetSheetIndex, id, index);

        public static Change WindowVisibleSheet(WindowId id, SheetId sheetId) =>
            new Change<WindowId, SheetId>(ChangeType.SetWindowVisibleSheet, id, sheetId);

        public static Change SheetTabColor(SheetId id, Color color) =>
            new Change<SheetId, Color>(ChangeType.SetSheetTabColor, id, color);        
    }
}
