using System.Windows;
using ExcelBrowser.Controller;
using ExcelBrowser.Model;
using System.ComponentModel;
using System.Windows.Controls;

namespace ExcelBrowser.UI {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public MainWindow() {
            InitializeComponent();

            this.monitor = new SessionMonitor(refreshSeconds: 0.05);
            txt_Session.DataContext = monitor;
        }

        private readonly SessionMonitor monitor;

        public TextBlock TextBlock => txt_Session;
    }
}
