using System.Collections.ObjectModel;

namespace ExcelBrowser.ViewModels {

    public class BookViewModel {

        public BookViewModel() {
            Sheets = new ObservableCollection<SheetViewModel>();
            Windows = new BookWindowsViewModel();
        }

        public string Name { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public bool IsAddIn { get; set; }
        public int WindowCount { get; set; }
        public ObservableCollection<SheetViewModel> Sheets { get; set; }
        public BookWindowsViewModel Windows { get; set; }
    }
}
