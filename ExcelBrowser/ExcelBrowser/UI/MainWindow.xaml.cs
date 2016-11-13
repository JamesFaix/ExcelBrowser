using System;
using System.Windows;
using ExcelBrowser.Controller;

namespace ExcelBrowser.UI {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable {

        public MainWindow(SessionMonitor monitor) {
            Requires.NotNull(monitor, nameof(monitor));
            InitializeComponent();

            this.monitor = monitor;            
        }

        private readonly SessionMonitor monitor;

        public void Dispose() {
            monitor.Dispose();
        }
    }
}
