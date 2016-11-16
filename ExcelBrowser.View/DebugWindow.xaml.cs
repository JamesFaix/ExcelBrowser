using System;
using System.Windows;
using ExcelBrowser.Controller;
using System.Diagnostics.CodeAnalysis;

namespace ExcelBrowser.View {

    /// <summary>
    /// Interaction logic for DebugWindow.xaml
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]
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

        [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]
        public void Dispose() {
            monitor.Dispose();
        }
    }
}
