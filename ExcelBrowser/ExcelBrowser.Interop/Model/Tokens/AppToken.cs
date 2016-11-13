using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using ExcelBrowser.Interop;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using xlBook = Microsoft.Office.Interop.Excel.Workbook;
using xlWin = Microsoft.Office.Interop.Excel.Window;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    /// <summary>
    /// Represents a snapshot of an Excel application instance.
    /// </summary>
    [DataContract]
    public class AppToken : Token<AppId> {

        public AppToken(xlApp app) : this(app?.Id()) {
            try {
                IsVisible = app.Visible
                    && app.AsProcess().IsVisible();
            }
            catch (COMException x)
            when (x.Message.StartsWith("The message filter indicated that the application is busy.")) {
                //This means the application is in a state that does not permit COM automation.
                //Often, this is due to a dialog window or right-click context menu being open.
                Debug.WriteLine($"Busy @ {Id}");
                IsVisible = false;
            }

            if (IsVisible) {
                Books = app.Workbooks.OfType<xlBook>()
                    .Select(wb => new BookToken(wb))
                    .ToImmutableArray();

                xlBook activeBook = app.ActiveWorkbook;
                if (activeBook != null) {
                    var id = activeBook.Id();
                    ActiveBook = Books.Single(b => Equals(b.Id, id));
                }

                xlWin activeWindow = app.ActiveWindow;
                if (activeWindow != null) {
                    var id = activeWindow.Id();
                    ActiveWindow = Books.Single(b => Equals(b.Id.BookName, id.BookName))
                        .Windows.Single(w => Equals(w.Id, id));
                }
            }
        }

        private AppToken(AppId id) : base(id) { }

        public static AppToken Unreachable(int processId) {
            return new AppToken(new AppId(processId)) {
                Books = new BookToken[0].ToImmutableArray()
            };
        }

        [DataMember(Order = 1)]
        public bool IsVisible { get; }

        [DataMember(Order = 2)]
        public IEnumerable<BookToken> Books { get; private set; }

        [DataMember(Order = 3)]
        public BookToken ActiveBook { get; }

        [DataMember(Order = 4)]
        public WindowToken ActiveWindow { get; }

        #region Equality

        public bool Equals(AppToken other) => base.Equals(other)
            && IsVisible == other.IsVisible
            && Books.SequenceEqual(other.Books)
            && Equals(ActiveBook, other.ActiveBook)
            && Equals(ActiveWindow, other.ActiveWindow);

        public override bool Equals(object obj) => Equals(obj as AppToken);

        public bool Matches(AppToken other) => base.Equals(other);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
