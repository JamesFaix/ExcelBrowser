using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Linq;
using ExcelBrowser.Model;
using ExcelBrowser.Monitoring;
using MoreLinq;
using static ExcelBrowser.Model.ModelChangeType;

namespace ExcelBrowser.Controller {

    public static class SessionChangeAnalyzer {

        public static IEnumerable<ModelChange> FindChanges(ValueChange<SessionToken> sessionChange) {
            Requires.NotNull(sessionChange, nameof(sessionChange));
            // Debug.WriteLine("SessionChangeAnalayzer.FindChanges");

            //Check for new session
            if (sessionChange.OldValue == null) {
                return ImmutableArray.Create(ModelChange.Create(sessionChange.NewValue.Id, Add));
            }
            else {
                return GetSessionChanges(sessionChange).ToImmutableArray();
            }
        }

        private static IEnumerable<ModelChange> GetSessionChanges(ValueChange<SessionToken> change) {
            //  Debug.WriteLine($"SessionChangeAnalyzer.GetSessionChanges({change})");
            return GetChanges<SessionId, SessionToken, AppId, AppToken>(change, (session => session.Apps), GetAppChanges)
                 .Concat(GetIfIdChanged<SessionId, SessionToken, AppId, AppToken>(change, (app => app.ActiveApp), Activate));
        }

        private static IEnumerable<ModelChange> GetAppChanges(ValueChange<AppToken> change) {
            //  Debug.WriteLine($"SessionChangeAnalyzer.GetAppChanges({change})");
            return GetIfChanged(change, (a => a.IsReachable), change.NewValue.Id, SetReachabilty)
                .Concat(GetChanges<AppId, AppToken, BookId, BookToken>(change, (app => app.Books), GetBookChanges))
                .Concat(GetIfIdChanged<AppId, AppToken, BookId, BookToken>(change, (app => app.ActiveBook), Activate))
                .Concat(GetIfIdChanged<AppId, AppToken, WindowId, WindowToken>(change, (app => app.ActiveWindow), Activate));
        }

        private static IEnumerable<ModelChange> GetBookChanges(ValueChange<BookToken> change) {
            //   Debug.WriteLine($"SessionChangeAnalyzer.GetBookChanges({change})");
            return GetSheetChanges(change)
                .Concat(GetWindowChanges(change));
        }

        private static IEnumerable<ModelChange> GetSheetChanges(ValueChange<BookToken> change) {
            //  Debug.WriteLine($"SessionChangeAnalyzer.GetSheetChanges({change})");
            return GetChanges<BookId, BookToken, SheetId, SheetToken>(change, (book => book.Sheets))
                .Concat(GetIfIdChanged<BookId, BookToken, SheetId, SheetToken>(change, (book => book.ActiveSheet), Activate));
        }

        private static IEnumerable<ModelChange> GetWindowChanges(ValueChange<BookToken> change) {
            //   Debug.WriteLine($"SessionChangeAnalyzer.GetWindowChanges({change})");
            return GetChanges<BookId, BookToken, WindowId, WindowToken>(change, (book => book.Windows));
        }

        #region Implementation

        private static IEnumerable<ModelChange> GetChanges<TParentId, TParentToken, TChildId, TChildToken>(
            ValueChange<TParentToken> change,
            Func<TParentToken, IEnumerable<TChildToken>> getChildren,
            Func<ValueChange<TChildToken>, IEnumerable<ModelChange>> drillDown = null)
            where TParentToken : Token<TParentId>
            where TChildToken : Token<TChildId> {
            // Debug.WriteLine($"SessionChangeAnalyzer.GetChanges({change})");
            Debug.Assert(Equals(change.OldValue.Id, change.NewValue.Id));

            var children = change.Select(getChildren);
            var ids = IdChanges.Create<TChildId, TChildToken>(children);

            var result = ids.Removed.Select(id => ModelChange.Create(id, Remove))
                .Concat(ids.Added.Select(id => ModelChange.Create(id, Add)));

            if (drillDown != null)
                result = result.Concat(ids.Changes.SelectMany(drillDown));

            return result;
        }

        private static IEnumerable<ModelChange> GetIfIdChanged<TParentId, TParentToken, TChildId, TChildToken>(
            ValueChange<TParentToken> change,
            Func<TParentToken, TChildToken> selector, ModelChangeType changeType)
            where TParentToken : Token<TParentId>
            where TChildToken : Token<TChildId> {

            var ids = change.Select(v => {
                var key = selector(v);
                return (key == null) ? default(TChildId) : key.Id;
            });

            if (ids.IsDifferent()) {
                var id = ids.NewValue;
                if (Equals(id, null)) id = ids.OldValue;

                return new[] { ModelChange.Create(id, changeType) };
            }
            else {
                return new ModelChange[0];
            }
        }
        
        private static IEnumerable<ModelChange> GetIfChanged<TSource, TKey, TId>(ValueChange<TSource> change, Func<TSource, TKey> selector, TId id, ModelChangeType changeType) {

            var keys = change.Select(selector);

            return keys.IsDifferent()
                ? new[] { ModelChange.Create(id, changeType) }
                : new ModelChange[0];
        }

        #endregion
    }

    internal class IdChanges<TId, TToken>
        where TToken : Token<TId> {

        public IdChanges(ValueChange<IEnumerable<TToken>> tokens) {

            Old = tokens.OldValue.Ids().ToArray();
            New = tokens.NewValue.Ids().ToArray();

            Removed = tokens.OldValue.ExceptIds(New).Ids().ToArray();
            Added = tokens.NewValue.ExceptIds(Old).Ids().ToArray();
            Persistent = New.Intersect(Old).ToArray();

            var persistedOld = tokens.OldValue.IntersectIds(Persistent).OrderBy(t => t.Id);
            var persistedNew = tokens.NewValue.IntersectIds(Persistent).OrderBy(t => t.Id);

            var persistedPairs = persistedOld.Zip(persistedNew, ValueChange.Create);

            Changes = persistedPairs.Where(vc => vc.IsDifferent()).ToArray();
        }

        public TId[] Old { get; }
        public TId[] New { get; }
        public TId[] Added { get; }
        public TId[] Removed { get; }
        public TId[] Persistent { get; }
        public ValueChange<TToken>[] Changes { get; }
    }

    internal static class IdChanges {

        public static IdChanges<TId, TToken> Create<TId, TToken>(ValueChange<IEnumerable<TToken>> tokens)
            where TToken : Token<TId> =>
            new IdChanges<TId, TToken>(tokens);

        public static IEnumerable<TId> Ids<TId>(this IEnumerable<Token<TId>> tokens) =>
           tokens.Select(t => t.Id);

        public static IEnumerable<TToken> ExceptIds<TId, TToken>(this IEnumerable<TToken> tokens, IEnumerable<TId> ids)
            where TToken : Token<TId> =>
            tokens.Where(t => !ids.Contains(t.Id));

        public static IEnumerable<TToken> IntersectIds<TId, TToken>(this IEnumerable<TToken> tokens, IEnumerable<TId> ids)
           where TToken : Token<TId> =>
           tokens.Where(t => ids.Contains(t.Id));
    }
}
