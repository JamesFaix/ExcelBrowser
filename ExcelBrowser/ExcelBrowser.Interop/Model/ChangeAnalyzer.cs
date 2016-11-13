using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using ExcelBrowser.Monitoring;
using MoreLinq;

namespace ExcelBrowser.Model {

    internal static class ChangeAnalyzer {

        public static IEnumerable<Change> FindChanges(ValueChange<SessionToken> sessionChange) {
            Requires.NotNull(sessionChange, nameof(sessionChange));
            // Debug.WriteLine("SessionChangeAnalayzer.FindChanges");

            //Check for new session
            if (sessionChange.OldValue == null) {
                return ImmutableArray.Create(Change.SessionStart(sessionChange.NewValue.Id));
            }
            else {
                return GetSessionChanges(sessionChange).ToImmutableArray();
            }
        }

        private static IEnumerable<Change> GetSessionChanges(ValueChange<SessionToken> diff) {
            var ids = new ChangeSet<AppId, AppToken>(diff.Select(session => session.Apps));

            return ids.RemovedChanges
                 .Concat(ids.AddedChanges)
                 .Concat(ids.NestedChanges(GetAppChanges));
        }

        private static IEnumerable<Change> GetAppChanges(ValueChange<AppToken> diff) {

            var ids = new ChangeSet<BookId, BookToken>(diff.Select(app => app.Books));

            var result = Enumerable.Empty<Change>();

            if (diff.IsChanged(app => app.IsActive))
                result = result.Concat(Change.SetActive(diff.NewValue.Id, diff.NewValue.IsActive));

            if (diff.IsChanged(app => app.IsVisible))
                result = result.Concat(Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible));

            result = result.Concat(ids.RemovedChanges)
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetBookChanges));

            return result;
        }

        private static IEnumerable<Change> GetBookChanges(ValueChange<BookToken> diff) {
            //   Debug.WriteLine($"SessionChangeAnalyzer.GetBookChanges({change})");

            var result = Enumerable.Empty<Change>();

            if (diff.IsChanged(b => b.IsActive))
                result = result.Concat(Change.SetActive(diff.NewValue.Id, diff.NewValue.IsActive));

            if (diff.IsChanged(b => b.IsVisible))
                result = result.Concat(Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible));

            result = result.Concat(GetSheetCollectionChanges(diff)
                .Concat(GetWindowCollectionChanges(diff)));

            return result;
        }

        private static IEnumerable<Change> GetSheetCollectionChanges(ValueChange<BookToken> diff) {
            var ids = new ChangeSet<SheetId, SheetToken>(diff.Select(book => book.Sheets));

            return ids.RemovedChanges
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetSheetChanges));
        }

        private static IEnumerable<Change> GetSheetChanges(ValueChange<SheetToken> diff) {
            if (diff.IsChanged(s => s.IsActive))
                yield return Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsActive);

            if (diff.IsChanged(s => s.IsVisible))
                yield return Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible);

            if (diff.IsChanged(s => s.Index))
                yield return Change.SheetMove(diff.NewValue.Id, diff.NewValue.Index);
        }

        private static IEnumerable<Change> GetWindowCollectionChanges(ValueChange<BookToken> diff) {
            var ids = new ChangeSet<WindowId, WindowToken>(diff.Select(book => book.Windows));

            return ids.RemovedChanges
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetWindowChanges));
        }

        private static IEnumerable<Change> GetWindowChanges(ValueChange<WindowToken> diff) {
            if (diff.IsChanged(s => s.IsActive))
                yield return Change.SetActive(diff.NewValue.Id, diff.NewValue.IsActive);

            if (diff.IsChanged(s => s.State))
                yield return Change.WindowSetState(diff.NewValue.Id, diff.NewValue.State);

            if (diff.IsChanged(s => s.IsVisible))
                yield return Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible);
        }
    }
}
