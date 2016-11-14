using System.Collections.ObjectModel;

namespace ExcelBrowser.ViewModels {

    public class SessionViewModel {

        public SessionViewModel() {
            Apps = new ObservableCollection<AppViewModel>();
        }

        public ObservableCollection<AppViewModel> Apps { get; set; }
    }
}
