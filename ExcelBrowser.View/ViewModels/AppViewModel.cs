using System.Collections.ObjectModel;

namespace ExcelBrowser.ViewModels {

    public class AppViewModel {

        public AppViewModel() {
            Books = new ObservableCollection<BookViewModel>();
        }

        public int ProcessId { get; set; }
        public string Version { get; set; }

        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }

        public string Label => $"{Version} [ProcessID: {ProcessId}]";

        public ObservableCollection<BookViewModel> Books { get; set; }
    }
}
