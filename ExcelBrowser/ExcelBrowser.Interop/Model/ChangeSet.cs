using System;
using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Monitoring;

namespace ExcelBrowser.Model {

    internal class ChangeSet<TId, TToken>
       where TToken : Token<TId> {

        public ChangeSet(ValueChange<IEnumerable<TToken>> tokens) {

            OldIds = tokens.OldValue.Ids().ToArray();
            NewIds = tokens.NewValue.Ids().ToArray();

            RemovedIds = tokens.OldValue.ExceptIds(NewIds).Ids().ToArray();
            AddedIds = tokens.NewValue.ExceptIds(OldIds).Ids().ToArray();
            PersistentIds = NewIds.Intersect(OldIds).ToArray();

            var persistedOld = tokens.OldValue.IntersectIds(PersistentIds).OrderBy(t => t.Id);
            var persistedNew = tokens.NewValue.IntersectIds(PersistentIds).OrderBy(t => t.Id);

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
    }
}
