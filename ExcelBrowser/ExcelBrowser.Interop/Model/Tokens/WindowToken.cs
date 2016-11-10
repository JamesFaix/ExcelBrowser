using System;
using xlWin = Microsoft.Office.Interop.Excel.Window;
using ExcelBrowser.Interop;
using System.Diagnostics;

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of an Excel window.
    /// </summary>
    public class WindowToken : Token<WindowId>, IEquatable<WindowToken> {

        public WindowToken(xlWin window) : base(window?.Id()) {
          //  Debug.WriteLine("WindowToken.Constructor");
            State = window.WindowState.Outer();
        }

        public WindowState State { get; }

        public bool Equals(WindowToken other) => base.Equals(other);

        public override bool Equals(object obj) => Equals(obj as WindowToken);
    }
}
