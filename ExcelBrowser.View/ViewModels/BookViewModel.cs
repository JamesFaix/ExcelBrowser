using System.Collections.ObjectModel;
using System.Windows.Media;
using System.Windows;

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
        
        public Brush Foreground => new SolidColorBrush { Color = Colors.Black };
        public Brush Background => new SolidColorBrush { Color = IsActive ? Colors.White : Colors.LightGray };
        public FontWeight FontWeight => IsActive ? FontWeights.Bold : FontWeights.Normal;
    }
}
