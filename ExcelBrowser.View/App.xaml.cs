using System.Windows;
using ExcelBrowser.Controller;
using ExcelBrowser.View;

namespace ExcelBrowser {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {

        protected override void OnStartup(StartupEventArgs e) {
            base.OnStartup(e);

            var monitor = new SessionMonitor(refreshSeconds:1);

            var window = new MainWindow(monitor);
            window.Show();
            
            var debugWindow = new DebugWindow(monitor);
            debugWindow.Show();
        }
    }
}
