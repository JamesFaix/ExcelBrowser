using System;
using System.Windows;
using ExcelBrowser.Controller;

namespace ExcelBrowser.View {
    /// <summary>
    /// Interaction logic for DebugWindow.xaml
    /// </summary>
    public partial class DebugWindow : Window, IDisposable {

        public DebugWindow(SessionMonitor monitor) {
            Requires.NotNull(monitor, nameof(monitor));
            InitializeComponent();

            this.monitor = monitor;
            this.log = new SessionLog(monitor);

            txt_Session.DataContext = monitor;
            txt_Log.DataContext = log;
        }

        private readonly SessionMonitor monitor;
        private readonly SessionLog log;       

        public void Dispose() {
            monitor.Dispose();
        }
    }
}
