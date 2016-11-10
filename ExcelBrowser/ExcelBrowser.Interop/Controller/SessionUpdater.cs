using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelBrowser.Model;
using ExcelBrowser.Monitoring;
using ExcelBrowser.Interop;

namespace ExcelBrowser.Controller {

    public class SessionUpdater : IDisposable {

        public SessionUpdater(double refreshSeconds = 0.05) {
            Requires.Positive(refreshSeconds, nameof(refreshSeconds));
            Debug.WriteLine("SessionUpdater.Constructor");

            this.session = Session.Current;
            this.SessionToken = new SessionToken(session);

            this.sessionMonitor = new ChangeDetector<SessionToken>(
                getValue: () => new SessionToken(this.session),
                refreshSeconds: refreshSeconds);

            sessionMonitor.Changed += SessionChanged;
        }

        private readonly Session session;
        private readonly ChangeDetector<SessionToken> sessionMonitor;
        public SessionToken SessionToken { get; private set; }

        private void SessionChanged(object sender, EventArgs<ValueChange<SessionToken>> e) {
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
            sessionMonitor.Dispose();
        }
    }
}
