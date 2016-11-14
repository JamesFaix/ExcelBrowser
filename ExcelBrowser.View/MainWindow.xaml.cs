using System;
using System.Windows;
using ExcelBrowser.Controller;
using ExcelBrowser.ViewModels;

namespace ExcelBrowser.View {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable {

        public MainWindow(SessionMonitor monitor) {
            Requires.NotNull(monitor, nameof(monitor));
            InitializeComponent();

            var viewModelUpdater = new ViewModelUpdater(monitor);
            viewModelUpdater.PropertyChanged += (sender, e) =>
            Dispatcher.Invoke(() => {
                sessionPanel.DataContext = viewModelUpdater.ViewModel;
            });
            //            ctrl_Session.DataContext = viewModelUpdater.ViewModel;

            this.monitor = monitor;
        }

        private readonly SessionMonitor monitor;

        public void Dispose() {
            monitor.Dispose();
        }
    }
}
