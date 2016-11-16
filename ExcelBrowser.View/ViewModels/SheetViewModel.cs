using System.Collections.ObjectModel;
using System.Windows.Media;
using System.Windows;
using ExcelBrowser.Model;

namespace ExcelBrowser.ViewModels {

    public class SheetViewModel {

        public SheetViewModel(SheetToken token) {
            Requires.NotNull(token, nameof(token));
            Token = token;
            Windows = new ObservableCollection<SheetWindowViewModel>();
        }

        public SheetToken Token { get; }

        public ObservableCollection<SheetWindowViewModel> Windows { get; set; }

        public string Label => Token.Id.SheetName;

        public FontWeight FontWeight => Token.IsActive ? FontWeights.Bold : FontWeights.Normal;

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
