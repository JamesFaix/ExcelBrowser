using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelBrowser.Model;
using ExcelBrowser.Monitoring;
using ExcelBrowser.Interop;

namespace ExcelBrowser.Controller {

    public sealed class SessionUpdater : IDisposable {

        public SessionUpdater(double refreshSeconds = 0.05) {
            Requires.Positive(refreshSeconds, nameof(refreshSeconds));
            Debug.WriteLine("SessionUpdater.Constructor");

            this.session = Session.Current;
            this.SessionToken = new SessionToken(session);

            this.detector = new ChangeDetector<SessionToken>(
                getValue: () => new SessionToken(this.session),
                refreshSeconds: refreshSeconds);

            detector.Changed += DetectorChanged;
        }

        private readonly Session session;
        private readonly ChangeDetector<SessionToken> detector;
        public SessionToken SessionToken { get; private set; }

        private void DetectorChanged(object sender, EventArgs<ValueChange<SessionToken>> e) {
            Debug.WriteLine("SessionUpdater.SessionChanged");
            var change = e.Value;
            var modelChanges = SessionChangeAnalyzer.FindChanges(change);
            if (modelChanges.Any()) {
                SessionToken = change.NewValue;
                OnChanged(modelChanges);
            }
        }

        public event EventHandler<EventArgs<IEnumerable<ModelChange>>> Changed;
        private void OnChanged(IEnumerable<ModelChange> changes) => 
            Changed?.Invoke(this, new EventArgs<IEnumerable<ModelChange>>(changes));

        public void Dispose() {
            detector.Dispose();
        }
    }
}
