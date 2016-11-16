using System;
using System.Windows;
using ExcelBrowser.Controller;
using ExcelBrowser.ViewModels;
using System.Diagnostics.CodeAnalysis;

namespace ExcelBrowser.View {

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]
    public partial class MainWindow : Window, IDisposable {

        public MainWindow(SessionMonitor monitor) {
            Requires.NotNull(monitor, nameof(monitor));
            InitializeComponent();

            this.monitor = monitor;
            this.viewModelUpdater = new ViewModelUpdater(monitor);

            viewModelUpdater.PropertyChanged += (sender, e) => SessionChanged();
            //Must wait for monitors refresh to happen          
            SessionChanged();
        }

        private readonly SessionMonitor monitor;
        private readonly ViewModelUpdater viewModelUpdater;

        private void SessionChanged() {
            Dispatcher.Invoke(() => {
                sessionPanel.DataContext = viewModelUpdater.ViewModel;
            });
        }

        [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]
        public void Dispose() {
            monitor.Dispose();
        }
    }
}
