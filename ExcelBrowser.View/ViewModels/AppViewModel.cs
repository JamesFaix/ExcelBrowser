using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;

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
        
        public Brush Foreground => new SolidColorBrush { Color = Colors.White };
        public Brush Background => new SolidColorBrush { Color = Colors.DarkGreen };
        public FontWeight FontWeight => IsActive ? FontWeights.Bold : FontWeights.Normal;
    }
}
