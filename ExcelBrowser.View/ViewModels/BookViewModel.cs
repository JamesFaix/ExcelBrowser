using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using ExcelBrowser.Model;
using System.Windows.Input;
using ExcelBrowser.Controller;

namespace ExcelBrowser.ViewModels {

    public class BookViewModel {

        public BookViewModel(BookToken token) {
            Requires.NotNull(token, nameof(token));
            Token = token;
            Sheets = new ObservableCollection<SheetViewModel>();
            Windows = new BookWindowsViewModel();

            Activate = new RelayCommand(obj => SessionCommands.ActivateBook(Token.Id));
        }

        public BookToken Token { get; }

        public ObservableCollection<SheetViewModel> Sheets { get; set; }
        public BookWindowsViewModel Windows { get; set; }

        public string Label => Token.Id.BookName;

        public Brush Foreground => new SolidColorBrush { Color = Colors.Black };
        public Brush Background => new SolidColorBrush { Color = Token.IsActive ? Colors.White : Colors.LightGray };
        public FontWeight FontWeight => Token.IsActive ? FontWeights.Bold : FontWeights.Normal;

        public ICommand Activate { get; }
    }
}
