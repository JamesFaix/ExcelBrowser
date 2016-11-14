using System;
using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Monitoring;

namespace ExcelBrowser.Model {

    internal class ChangeSet<TParentId, TParentToken, TChildId, TChildToken>
       where TParentToken : Token<TParentId>
       where TChildToken : Token<TChildId> {

        public ChangeSet(ValueChange<TParentToken> parentToken, ValueChange<IEnumerable<TChildToken>> childTokens) {

            OldIds = Ids(childTokens.OldValue).ToArray();
            NewIds = Ids(childTokens.NewValue).ToArray();

            RemovedIds = Ids(ExceptIds(childTokens.OldValue, NewIds)).ToArray();
            AddedIds = Ids(ExceptIds(childTokens.NewValue, OldIds)).ToArray();
            PersistentIds = NewIds.Intersect(OldIds).ToArray();

            var persistedOld = IntersectIds(childTokens.OldValue, PersistentIds).OrderBy(t => t.Id);
            var persistedNew = IntersectIds(childTokens.NewValue, PersistentIds).OrderBy(t => t.Id);
            var persistedPairs = persistedOld.Zip(persistedNew, ValueChange.Create);
            Diffs = persistedPairs.Where(vc => vc.IsChanged()).ToArray();

            AddedChanges = AddedIds.Select(id => Change.Added(parentToken.NewValue.Id, id)).ToArray();
            RemovedChanges = RemovedIds.Select(id => Change.Removed(parentToken.NewValue.Id, id)).ToArray();
        }

        public IEnumerable<TChildId> OldIds { get; }
        public IEnumerable<TChildId> NewIds { get; }
        public IEnumerable<TChildId> AddedIds { get; }
        public IEnumerable<TChildId> RemovedIds { get; }
        public IEnumerable<TChildId> PersistentIds { get; }
        public IEnumerable<ValueChange<TChildToken>> Diffs { get; }

        public IEnumerable<Change> AddedChanges { get; }
        public IEnumerable<Change> RemovedChanges { get; }

        public IEnumerable<Change> NestedChanges(Func<ValueChange<TChildToken>, IEnumerable<Change>> selector) {
            Requires.NotNull(selector, nameof(selector));
            return Diffs.SelectMany(selector);
        }

        private IEnumerable<TChildId> Ids(IEnumerable<Token<TChildId>> tokens) =>
            tokens.Select(t => t.Id);

        private IEnumerable<TChildToken> ExceptIds(IEnumerable<TChildToken> tokens, IEnumerable<TChildId> ids) =>
            tokens.Where(t => !ids.Contains(t.Id));

        private IEnumerable<TChildToken> IntersectIds(IEnumerable<TChildToken> tokens, IEnumerable<TChildId> ids) =>
            tokens.Where(t => ids.Contains(t.Id));
    }
}
