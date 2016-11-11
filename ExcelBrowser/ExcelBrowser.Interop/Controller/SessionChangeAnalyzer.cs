using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using ExcelBrowser.Model;
using ExcelBrowser.Monitoring;
using MoreLinq;

namespace ExcelBrowser.Controller {

    public static class SessionChangeAnalyzer {

        public static IEnumerable<ModelChange> FindChanges(ValueChange<SessionToken> sessionChange) {
            Requires.NotNull(sessionChange, nameof(sessionChange));
            // Debug.WriteLine("SessionChangeAnalayzer.FindChanges");

            //Check for new session
            if (sessionChange.OldValue == null) {
                return ImmutableArray.Create(ModelChange.SessionStart(sessionChange.NewValue.Id));
            }
            else {
                return GetSessionChanges(sessionChange).ToImmutableArray();
            }
        }

        private static IEnumerable<ModelChange> GetSessionChanges(ValueChange<SessionToken> diff) {

            var ids = new ChangeSet<AppId, AppToken>(diff.Select(session => session.Apps));

            var result = ids.Removes
                .Concat(ids.Adds)
                .Concat(ids.NestedChanges(GetAppChanges));

            if (diff.Select(session => session.ActiveApp.Id).IsDifferent()) {
                result = result.Concat(ModelChange.Activated(diff.NewValue.ActiveApp.Id));
            }
            return result;
        }

        private static IEnumerable<ModelChange> GetAppChanges(ValueChange<AppToken> diff) {

            var ids = new ChangeSet<BookId, BookToken>(diff.Select(app => app.Books));

            var result = Enumerable.Empty<ModelChange>();

            if (diff.Select(app => app.IsReachable).IsDifferent())
                result = result.Concat(ModelChange.AppSetReachablity(diff.NewValue.Id, diff.NewValue.IsReachable));

            result = result.Concat(ids.Removes)
                .Concat(ids.Adds)
                .Concat(ids.NestedChanges(GetBookChanges));

            if (diff.Select(app => app.ActiveBook?.Id).IsDifferent())
                result = result.Concat(ModelChange.Activated(diff.NewValue.ActiveBook.Id));

            if (diff.Select(app => app.ActiveWindow?.Id).IsDifferent())
                result = result.Concat(ModelChange.Activated(diff.NewValue.ActiveWindow.Id));

            return result;
        }

        private static IEnumerable<ModelChange> GetBookChanges(ValueChange<BookToken> diff) {
            //   Debug.WriteLine($"SessionChangeAnalyzer.GetBookChanges({change})");

            //Book visibility changes

            var result = GetSheetChanges(diff)
                .Concat(GetWindowChanges(diff));

            if (diff.Select(book => book.ActiveSheet?.Id).IsDifferent()) {
                result = result.Concat(ModelChange.Activated(diff.NewValue.ActiveSheet.Id));
            }

            return result;
        }

        private static IEnumerable<ModelChange> GetSheetChanges(ValueChange<BookToken> diff) {

            var ids = new ChangeSet<SheetId, SheetToken>(diff.Select(book => book.Sheets));

            var result = ids.Removes
                .Concat(ids.Adds);

            //    .Concat(ids.NestedChanges(GetAppChanges));

            return result;
        }

        private static IEnumerable<ModelChange> GetWindowChanges(ValueChange<BookToken> diff) {

            var ids = new ChangeSet<WindowId, WindowToken>(diff.Select(book => book.Windows));

            var result = ids.Removes.Concat(ids.Adds);

            //    .Concat(ids.NestedChanges(GetAppChanges));

            return result;
        }
    }
}
