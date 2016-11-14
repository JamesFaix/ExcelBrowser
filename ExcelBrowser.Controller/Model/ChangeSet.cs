using System;
using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Monitoring;

namespace ExcelBrowser.Model {

    internal class ChangeSet<TId, TToken>
       where TToken : Token<TId> {

        public ChangeSet(ValueChange<IEnumerable<TToken>> tokens) {

            OldIds = Ids(tokens.OldValue).ToArray();
            NewIds = Ids(tokens.NewValue).ToArray();

            RemovedIds = Ids(ExceptIds(tokens.OldValue, NewIds)).ToArray();
            AddedIds = Ids(ExceptIds(tokens.NewValue, OldIds)).ToArray();
            PersistentIds = NewIds.Intersect(OldIds).ToArray();

            var persistedOld = IntersectIds(tokens.OldValue, PersistentIds).OrderBy(t => t.Id);
            var persistedNew = IntersectIds(tokens.NewValue, PersistentIds).OrderBy(t => t.Id);

            var persistedPairs = persistedOld.Zip(persistedNew, ValueChange.Create);

            Diffs = persistedPairs.Where(vc => vc.IsChanged()).ToArray();
        }

        public TId[] OldIds { get; }
        public TId[] NewIds { get; }
        public TId[] AddedIds { get; }
        public TId[] RemovedIds { get; }
        public TId[] PersistentIds { get; }
        public ValueChange<TToken>[] Diffs { get; }

        public IEnumerable<Change> AddedChanges => AddedIds.Select(id => Change.Added(id));
        public IEnumerable<Change> RemovedChanges => RemovedIds.Select(id => Change.Removed(id));

        public IEnumerable<Change> NestedChanges(Func<ValueChange<TToken>, IEnumerable<Change>> selector) {
            Requires.NotNull(selector, nameof(selector));
            return Diffs.SelectMany(selector);
        }

        private IEnumerable<TId> Ids(IEnumerable<Token<TId>> tokens) =>
            tokens.Select(t => t.Id);

        private IEnumerable<TToken> ExceptIds(IEnumerable<TToken> tokens, IEnumerable<TId> ids) =>
            tokens.Where(t => !ids.Contains(t.Id));

        private IEnumerable<TToken> IntersectIds(IEnumerable<TToken> tokens, IEnumerable<TId> ids) =>
            tokens.Where(t => ids.Contains(t.Id));
    }
}
