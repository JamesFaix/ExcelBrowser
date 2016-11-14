using System.Windows.Controls;
using ExcelBrowser.ViewModels;

namespace ExcelBrowser.View.Controls {

    /// <summary>
    /// Interaction logic for SessionPanel.xaml
    /// </summary>
    public partial class SessionPanel : UserControl {

        public SessionPanel() {
            InitializeComponent();
        }

        private SessionViewModel viewModel;
        public SessionViewModel ViewModel {
            get { return viewModel; }
            set {
                viewModel = value;
                this.DataContext = value;
            }
        }
    }
}
