using System;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using xlWin = Microsoft.Office.Interop.Excel.Window;

namespace ExcelBrowser.Interop {

    /// <summary>
    /// Provides methods for creating objects or querying data from window handles.
    /// </summary>
    internal static class WindowHandleUtil {
        
        public static int GetProcessId(int windowHandle) {
            Requires.NotDefault(windowHandle, nameof(windowHandle));

            int processId;
            NativeMethods.GetWindowThreadProcessId(windowHandle, out processId);
            return processId;
        }       
         
        #region Excel App from Main Window handle

        public static xlApp GetAppFromMainHandle(int mainWindowHandle) {
            Requires.NotDefault(mainWindowHandle, nameof(mainWindowHandle));

            int childHandle = 0;
            NativeMethods.EnumChildWindows(mainWindowHandle, NextChildWindowHandle, ref childHandle);

            object obj;
            NativeMethods.AccessibleObjectFromWindow(childHandle, windowObjectId, windowInterfaceId, out obj);
            
            return (obj as xlWin)?.Application;
        }

        const uint windowObjectId = 0xFFFFFFF0;
        private static byte[] windowInterfaceId = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();

        private static bool NextChildWindowHandle(int currentChildHandle, ref int nextChildHandle) {
          //  Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - NextChildWindowHandle(" + currentChildHandle + ")");
            const string excelClassName = "EXCEL7";

            var result = true;

            var className = NativeMethods.GetClassName(currentChildHandle);
           // Debug.WriteLine(currentChildHandle + " ClassName: " + className);
            if (className == excelClassName) {
                nextChildHandle = currentChildHandle;
                result = false;
            }
          //  Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - NextChildWindowHandle(" + currentChildHandle + ", ref " + nextChildHandle + ") => " + result);
            return result;
        }
        
        #endregion
        
        public static int GetWindowZ(IntPtr windowHandle) {
            /// <summary>
            /// The retrieved handle identifies the window above the specified window in the Z order.
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window. 
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            const int GW_HWNDPREV = 3;

            var z = 0;
            //Count all windows above the starting window
            for (var h = windowHandle; h != IntPtr.Zero; h = NativeMethods.GetWindow(h, GW_HWNDPREV)) {
                z++;
            }
            return z;
        }
    }
}
