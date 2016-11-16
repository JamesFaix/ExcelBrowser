using System.Collections.ObjectModel;
using System.Windows.Media;
using System.Windows;

namespace ExcelBrowser.ViewModels {

    public class SheetViewModel {

        public SheetViewModel() {
            Windows = new ObservableCollection<SheetWindowViewModel>();
        }

        public string Name { get; set; }


        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public ObservableCollection<SheetWindowViewModel> Windows { get; set; }

        public FontWeight FontWeight => IsActive ? FontWeights.Bold : FontWeights.Normal;

        #region Color

        public Color? TabColor { get; set; }

        public Brush Foreground => new SolidColorBrush {
            Color = IsBackgroundDark ? Colors.White : Colors.Black
        };

        private bool IsBackgroundDark {
            get {
                if (TabColor == null) return false;

                var r = TabColor.Value.R;
                var g = TabColor.Value.G;
                var b = TabColor.Value.B;

                var average = (r + g + b) / 3;

                return average < 115;
            }
        }

        #endregion
    }
}
