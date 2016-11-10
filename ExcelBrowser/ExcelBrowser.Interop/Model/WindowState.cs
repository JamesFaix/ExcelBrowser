using xlState = Microsoft.Office.Interop.Excel.XlWindowState;

namespace ExcelBrowser.Model {

    public enum WindowState {
        Normal,
        Minimized,
        Maximized
    }

    public static class WindowStateExt {
        public static xlState Inner (this WindowState state) {
            switch (state) {
                case WindowState.Maximized: return xlState.xlMaximized;
                case WindowState.Minimized: return xlState.xlMinimized;
                case WindowState.Normal: return xlState.xlNormal;
                default: throw Requires.ValidEnum((int)state);
            }
        }

        public static WindowState Outer (this xlState state) {
            switch (state) {
                case xlState.xlMaximized: return WindowState.Maximized;
                case xlState.xlMinimized: return WindowState.Minimized; 
                case xlState.xlNormal:    return WindowState.Normal;
                default: throw Requires.ValidEnum((int)state);
            }
        }
    }
}
