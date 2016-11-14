using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using ExcelBrowser.Interop;
using ExcelBrowser.Model;
using ExcelBrowser.Monitoring;

namespace ExcelBrowser.Controller {

    public sealed class SessionMonitor : IDisposable, INotifyPropertyChanged {

        public SessionMonitor(double refreshSeconds = 0.05) {
            Requires.Positive(refreshSeconds, nameof(refreshSeconds));
            Debug.WriteLine($"{nameof(SessionMonitor)}.{nameof(SessionMonitor)}");

            this.session = Interop.Session.Current;
            this.sessionCache = new SessionCache(session);
            this.Session = sessionCache.Current;
            this.detector = new ChangeDetector<SessionToken>(
                getValue: () => sessionCache.Current,
                refreshSeconds: refreshSeconds);

            detector.Changed += DetectorChanged;
        }

        private readonly Session session;
        private readonly SessionCache sessionCache;
        private readonly ChangeDetector<SessionToken> detector;

        public SessionToken Session { get; private set; }

        public string SessionSerialized => Session.ToString();

        private void DetectorChanged(object sender, EventArgs<ValueChange<SessionToken>> e) {
            Debug.WriteLine($"{nameof(SessionMonitor)}.{nameof(DetectorChanged)}");
            var change = e.Value;
            var modelChanges = ChangeAnalyzer.FindChanges(change);
            if (modelChanges.Any()) {
                Session = change.NewValue;
                OnSessionChanged(modelChanges);
            }
        }

        public event EventHandler<EventArgs<IEnumerable<Change>>> SessionChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnSessionChanged(IEnumerable<Change> changes) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("SessionSerialized"));
            SessionChanged?.Invoke(this, new EventArgs<IEnumerable<Change>>(changes));
        }

        public void Dispose() {
            detector.Dispose();
        }
    }
}
