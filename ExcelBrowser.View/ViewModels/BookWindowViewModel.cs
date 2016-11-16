using System.Windows;
using System.Windows.Media;
using ExcelBrowser.Model;

namespace ExcelBrowser.ViewModels {

    public class BookWindowViewModel {

        public BookWindowViewModel(WindowToken token) {
            Requires.NotNull(token, nameof(token));
            Token = token;
        }

        public WindowToken Token { get; }

        public string Label => $"[{Token.Id.WindowIndex}]";

        public Brush Background => new SolidColorBrush { Color = Token.IsActive ? Colors.White : Colors.LightGray };
        public FontWeight FontWeight => Token.IsActive ? FontWeights.Bold : FontWeights.Normal;
    }
}
