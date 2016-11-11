using System;
using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Model;
using ExcelBrowser.Monitoring;

namespace ExcelBrowser.Controller {

    internal class ChangeSet<TId, TToken>
       where TToken : Token<TId> {

        public ChangeSet(ValueChange<IEnumerable<TToken>> tokens) {

            Old = tokens.OldValue.Ids().ToArray();
            New = tokens.NewValue.Ids().ToArray();

            Removed = tokens.OldValue.ExceptIds(New).Ids().ToArray();
            Added = tokens.NewValue.ExceptIds(Old).Ids().ToArray();
            Persistent = New.Intersect(Old).ToArray();

            var persistedOld = tokens.OldValue.IntersectIds(Persistent).OrderBy(t => t.Id);
            var persistedNew = tokens.NewValue.IntersectIds(Persistent).OrderBy(t => t.Id);

            var persistedPairs = persistedOld.Zip(persistedNew, ValueChange.Create);

            Diffs = persistedPairs.Where(vc => vc.IsDifferent()).ToArray();
        }

        public TId[] Old { get; }
        public TId[] New { get; }
        public TId[] Added { get; }
        public TId[] Removed { get; }
        public TId[] Persistent { get; }
        public ValueChange<TToken>[] Diffs { get; }

        public IEnumerable<ModelChange> Adds => Added.Select(id => ModelChange.Added(id));
        public IEnumerable<ModelChange> Removes => Removed.Select(id => ModelChange.Removed(id));

        public IEnumerable<ModelChange> NestedChanges(Func<ValueChange<TToken>, IEnumerable<ModelChange>> selector) {
            Requires.NotNull(selector, nameof(selector));
            return Diffs.SelectMany(selector);
        }
    }
}
