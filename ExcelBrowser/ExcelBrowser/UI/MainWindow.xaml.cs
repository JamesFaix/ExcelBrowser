using System;
using System.Collections.Generic;
using System.Windows;
using ExcelBrowser.Controller;
using ExcelBrowser.Model;

namespace ExcelBrowser.UI {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable {

        public MainWindow() {
            InitializeComponent();

            monitor = new SessionMonitor(refreshSeconds: 0.05);
            log = new SessionLog(monitor);

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
