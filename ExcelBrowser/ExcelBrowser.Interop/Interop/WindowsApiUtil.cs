using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using xlWin = Microsoft.Office.Interop.Excel.Window;

namespace ExcelBrowser.Interop {

    /// <summary>
    /// Encapsulates interop with native Windows API.
    /// </summary>
    static class WindowsApiUtil {

        private const string USER32 = "User32.dll";
        private const string OLEACC = "Oleacc.dll";

        #region Application to Process

        public static int GetProcessIdFromWindowHandle(int windowHandle) {
            Requires.NotDefault(windowHandle, nameof(windowHandle));

            int processId;
            GetWindowThreadProcessId(windowHandle, out processId);
            return processId;
        }

        /// <summary>Retrieves the identifier of the thread that created the specified window and optionally, 
        /// the identifier of the process that created the window.</summary>
        /// <param name="hWnd">A handle to the window.</param>
        /// <param name="lpdwProcessId">A pointer to a variable that receives the process identifier.
        /// If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process 
        /// to the variable; otherwise it does not.</param>
        /// <returns>The identifier of the thread that created the window.</returns>
        [DllImport(USER32)]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        #endregion

        #region WindowHandle to Excel Window
        public static xlWin ExcelWindowFromWindowHandle(int windowHandle) {
            Requires.NotDefault(windowHandle, nameof(windowHandle));

            xlWin window;
            AccessibleObjectFromWindow(windowHandle, windowObjectId, windowInterfaceId, out window);
            return window;
        }

        const uint windowObjectId = 0xFFFFFFF0;

        private static byte[] windowInterfaceId = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();

        /// <summary>Retrieves the address of the specified interface for the object associated with the specified window.</summary>
        /// <param name="hwnd">Specifies the handle of a window for which an object is to be retrieved. 
        /// To retrieve an interface pointer to the cursor or caret object, specify NULL and use the appropriate ID in dwObjectID.</param>
        /// <param name="dwObjectID">Specifies the object ID. This value is one of the standard object identifier constants or a custom object ID
        /// such as OBJID_NATIVEOM, which is the object ID for the Office native object model.</param>
        /// <param name="riid">Specifies the reference identifier of the requested interface. This value is either IID_IAccessible or IID_Dispatch,
        /// but it can also be IID_IUnknown, or the IID of any interface that the object is expected to support.</param>
        /// <param name="ppvObject">Address of a pointer variable that receives the address of the specified interface.</param>
        /// <returns>If successful, returns S_OK; otherwise returns E_INVALIDARG, E_NOINTERFACE, or another standard COM error code.</returns>
        [DllImport(OLEACC)]
        private static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, out xlWin ppvObject);

        #endregion

        #region Excel App from Main Window handle

        public static xlApp ExcelApplicationFromMainWindowHandle(int mainWindowHandle) {
            Requires.NotDefault(mainWindowHandle, nameof(mainWindowHandle));

            var childHandle = ChildHandleFromMainHandle(mainWindowHandle);
            xlWin window = ExcelWindowFromWindowHandle(childHandle);

            return window.Application;
        }

        private static int ChildHandleFromMainHandle(int mainHandle) {
        //    Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - ChildHandleFromMainHandle(" + mainHandle + ")");
            int childHandle = 0;
            EnumChildWindows(mainHandle, NextChildWindowHandle, ref childHandle);
            return childHandle;
        }

        /// <summary>Enumerates the child windows that belong to the specified parent window by passing the handle to each child window, in turn, 
        /// to an application-defined callback function. EnumChildWindows continues until the last child window is enumerated or 
        /// the callback function returns false.</summary>
        /// <param name="hWndParent">A handle to the parent window whose child windows are to be enumerated. If this parameter is NULL,
        /// this function is equivalent to EnumWindows.</param>
        /// <param name="lpEnumFunc">A point to an application-defined callback function.</param>
        /// <param name="lParam">An application-defined value to be passed tot he callback function.</param>
        /// <returns>The return value is not used.</returns>
        [DllImport(USER32)]
        private static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        /// <summary>An application-defined callback function used with the EnumChildWindows function. 
        /// It receives the child window handles.  The WNDENUMPROC type defines a pointer to this callback function.
        /// EnumChildProc is a placeholder for the application-defined function name.</summary>
        /// <param name="hwnd">A handle to the child window of the parent window specified in EnumChildWindows.</param>
        /// <param name="lParam">The application-defined value given in EnumChildWindows.</param>
        /// <returns>To continue enumeration, the callback function must return TRUE; to stop enumeration it must return FALSE.</returns>
        private delegate bool EnumChildCallback(int hwnd, ref int lParam);

        private static bool NextChildWindowHandle(int currentChildHandle, ref int nextChildHandle) {
          //  Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - NextChildWindowHandle(" + currentChildHandle + ")");
            const string excelClassName = "EXCEL7";

            var result = true;

            var className = GetClassName(currentChildHandle);
           // Debug.WriteLine(currentChildHandle + " ClassName: " + className);
            if (className == excelClassName) {
                nextChildHandle = currentChildHandle;
                result = false;
            }
          //  Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - NextChildWindowHandle(" + currentChildHandle + ", ref " + nextChildHandle + ") => " + result);
            return result;
        }

        //Gets the name of the COM class to which the specified window belongs.
        public static string GetClassName(int windowHandle) {
            var buffer = new StringBuilder(128);
            GetClassName(windowHandle, buffer, 128);
            var result = buffer.ToString();
            return result;
        }

        /// <summary>Retrieves the name of the class to which the specified window belongs.</summary>
        /// <param name="hWnd">A handle to the window and, indirectly, the class to which the window belongs.</param>
        /// <param name="lpClassName">The class name string.</param>
        /// <param name="nMaxCount">The length of the lpClassName buffer, in characters. The buffer must be large enough to include
        /// the terminating null character; otherwise, the class name string is truncated to nMaxCount-1 characters.</param>
        /// <returns>If the function succeeds, the number of characters copied to the buffer not including the terminating null character;
        /// otherwise 0.</returns>
        [DllImport(USER32)]
        private static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);

        #endregion

        #region Get Window Z

        public static int GetWindowZ(IntPtr windowHandle) {
            var z = 0;
            //Count all windows above the starting window
            for (var h = windowHandle; h != IntPtr.Zero; h = GetWindow(h, GW_HWNDPREV)) {
                z++;
            }
            return z;
        }

        /// <summary>
        /// The retrieved handle identifies the window above the specified window in the Z order.
        /// If the specified window is a topmost window, the handle identifies a topmost window.
        /// If the specified window is a top-level window, the handle identifies a top-level window. 
        /// If the specified window is a child window, the handle identifies a sibling window.
        /// </summary>
        const int GW_HWNDPREV = 3;

        /// <summary>Retrieves a handle to a window that has the specified relationship 
        /// (Z-Order or owner) to the specified window.</summary>
        /// <param name="hWnd">A handle to a window. The window handle retrieved is relative to this window, 
        /// based on the value of the uCmd parameter.</param>
        /// <param name="uCmd">The relationship between the specified window and the window whose handle is to be 
        /// retrieved. This parameter can be one of the following values. 
        /// (GW_CHILD, GW_ENABLEDPOPUP, GW_HWNDFIRST, GW_HWNDLAST, GW_HWNDNEXT, GWHWNDPREV, GW_OWNER)</param>
        /// <returns>If the function succeeds, a window handle; if no window exists with the specified relationship
        /// to the specified window, NULL.</returns>
        [DllImport("User32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        #endregion

    }
}
