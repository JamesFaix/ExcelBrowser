using System.Windows;
using ExcelBrowser.Controller;

namespace ExcelBrowser {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {

        protected override void OnStartup(StartupEventArgs e) {
            base.OnStartup(e);

            var monitor = new SessionMonitor(refreshSeconds:1);
            //var window = new View.DebugWindow(monitor);
            var window = new View.MainWindow(monitor);

            window.Show();
        }
    }
}
