﻿using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using ExcelBrowser.Monitoring;
using MoreLinq;

namespace ExcelBrowser.Model {

    public static class ChangeAnalyzer {

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
            var ids = new ChangeSet<SessionId, SessionToken, AppId, AppToken>(diff, diff.Select(session => session.Apps));

            return ids.RemovedChanges
                 .Concat(ids.AddedChanges)
                 .Concat(ids.NestedChanges(GetAppChanges));
        }

        private static IEnumerable<Change> GetAppChanges(ValueChange<AppToken> diff) {

            var ids = new ChangeSet<AppId, AppToken, BookId, BookToken>(diff, diff.Select(app => app.Books));

            var result = Enumerable.Empty<Change>();

            if (diff.IsChanged(app => app.IsVisible)) {
                var newValue = diff.NewValue.IsVisible;
                result = result.Concat(Change.SetVisibility(diff.NewValue.Id, newValue));
                if (!newValue) return result; //Don't look for other changes if the app just went invisible.
            }

            if (diff.IsChanged(app => app.IsActive))
                result = result.Concat(Change.SetActive(diff.NewValue.Id, diff.NewValue.IsActive));
            
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
            var ids = new ChangeSet<BookId, BookToken, SheetId, SheetToken>(diff, diff.Select(book => book.Sheets));

            return ids.RemovedChanges
                .Concat(ids.AddedChanges)
                .Concat(ids.NestedChanges(GetSheetChanges));
        }

        private static IEnumerable<Change> GetSheetChanges(ValueChange<SheetToken> diff) {
            if (diff.IsChanged(s => s.IsActive))
                yield return Change.SetActive(diff.NewValue.Id, diff.NewValue.IsActive);

            if (diff.IsChanged(s => s.IsVisible))
                yield return Change.SetVisibility(diff.NewValue.Id, diff.NewValue.IsVisible);

            if (diff.IsChanged(s => s.Index))
                yield return Change.SheetSetIndex(diff.NewValue.Id, diff.NewValue.Index);

            if (diff.IsChanged(s => s.TabColor))
                yield return Change.SheetTabColor(diff.NewValue.Id, diff.NewValue.TabColor);
        }

        private static IEnumerable<Change> GetWindowCollectionChanges(ValueChange<BookToken> diff) {
            var ids = new ChangeSet<BookId, BookToken, WindowId, WindowToken>(diff, diff.Select(book => book.Windows));

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

            if (diff.IsChanged(s => s.ActiveSheetId))
                yield return Change.WindowVisibleSheet(diff.NewValue.Id, diff.NewValue.ActiveSheetId);
        }
    }
}
