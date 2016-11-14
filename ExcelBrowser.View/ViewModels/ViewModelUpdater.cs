using System.Collections.Generic;
using ExcelBrowser.Controller;
using ExcelBrowser.Model;

namespace ExcelBrowser.ViewModels {

    public class ViewModelUpdater {

        public ViewModelUpdater() {

            ViewModel = new SessionPanelViewModel();

            this.monitor = new SessionMonitor(refreshSeconds: 0.05);
            this.monitor.SessionChanged += SessionChanged;
        }

        private readonly SessionMonitor monitor;

        public SessionPanelViewModel ViewModel { get; }

        private void SessionChanged(object sender, EventArgs<IEnumerable<Change>> e) {





        }
    }
}
