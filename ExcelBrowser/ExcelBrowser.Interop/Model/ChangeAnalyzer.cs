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

            var result = ids.RemovedChanges
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetAppChanges));

            if (diff.IsChanged(session => session.ActiveApp?.Id)) {
                var app = diff.NewValue.ActiveApp;
                if (app != null) {
                    result = result.Concat(Change.Activated(app.Id));
                }
            }
            return result;
        }

        private static IEnumerable<Change> GetAppChanges(ValueChange<AppToken> diff) {

            var ids = new ChangeSet<BookId, BookToken>(diff.Select(app => app.Books));

            var result = Enumerable.Empty<Change>();

            if (diff.IsChanged(app => app.IsVisible))
                result = result.Concat(Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible));

            result = result.Concat(ids.RemovedChanges)
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetBookChanges));

            if (diff.IsChanged(app => app.ActiveBook?.Id))
                result = result.Concat(Change.Activated(diff.NewValue.ActiveBook.Id));

            if (diff.IsChanged(app => app.ActiveWindow?.Id))
                result = result.Concat(Change.Activated(diff.NewValue.ActiveWindow.Id));

            return result;
        }

        private static IEnumerable<Change> GetBookChanges(ValueChange<BookToken> diff) {
            //   Debug.WriteLine($"SessionChangeAnalyzer.GetBookChanges({change})");

            //Book visibility changes

            var result = GetBookSheetChanges(diff)
                .Concat(GetBookWindowChanges(diff));

            if (diff.IsChanged(book => book.ActiveSheet?.Id)) {
                result = result.Concat(Change.Activated(diff.NewValue.ActiveSheet.Id));
            }

            return result;
        }

        private static IEnumerable<Change> GetBookSheetChanges(ValueChange<BookToken> diff) {

            var ids = new ChangeSet<SheetId, SheetToken>(diff.Select(book => book.Sheets));

            var result = ids.RemovedChanges
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetSingleSheetChanges));

            return result;
        }

        private static IEnumerable<Change> GetSingleSheetChanges(ValueChange<SheetToken> diff) {
            if (diff.IsChanged(s => s.IsVisible))
                yield return Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible);
            if (diff.IsChanged(s => s.Index))
                yield return Change.SheetMove(diff.NewValue.Id, diff.NewValue.Index);
        }

        private static IEnumerable<Change> GetBookWindowChanges(ValueChange<BookToken> diff) {

            var ids = new ChangeSet<WindowId, WindowToken>(diff.Select(book => book.Windows));

            var result = ids.RemovedChanges
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetSingleWindowChanges));

            return result;
        }

        private static IEnumerable<Change> GetSingleWindowChanges(ValueChange<WindowToken> diff) {
            if (diff.IsChanged(s => s.State))
                yield return Change.WindowSetState(diff.NewValue.Id, diff.NewValue.State);

            if (diff.IsChanged(s => s.IsVisible))
                yield return Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible);
        }
    }
}
