using System.Windows;
using System.Windows.Media;

namespace ExcelBrowser.ViewModels {

    public class BookWindowViewModel {

        public int Index { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }

        public string Label => $"[{Index}]";

        public Brush Background => new SolidColorBrush { Color = IsActive ? Colors.White : Colors.LightGray };
        public FontWeight FontWeight => IsActive ? FontWeights.Bold : FontWeights.Normal;
    }
}
