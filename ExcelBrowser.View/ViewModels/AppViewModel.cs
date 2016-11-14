using System.Collections.ObjectModel;

namespace ExcelBrowser.ViewModels {

    public class AppViewModel {

        public AppViewModel() {
            Books = new ObservableCollection<BookViewModel>();
        }

        public int ProcessId { get; set; }

        public string Label => $"ProcessID: {ProcessId}";

        public ObservableCollection<BookViewModel> Books { get; set; }
    }
}
