using System.Windows.Input;
using System.Windows.Media;
using ExcelBrowser.Controller;
using ExcelBrowser.Model;

namespace ExcelBrowser.ViewModels {

    public class SheetWindowViewModel {

        public SheetWindowViewModel(SheetToken token, int windowIndex) {
            Requires.NotNull(token, nameof(token));
            Token = token;
            WindowIndex = windowIndex;

            Activate = new RelayCommand(obj => SessionCommands.ActivateSheet(Token.Id, WindowIndex));
        }
        
        public SheetToken Token { get; }

        public int WindowIndex { get; }

        public bool IsActive { get; set; } //Not the same as Token.IsActive

        public Brush Background => new SolidColorBrush {
            Color = IsActive ? Colors.White : (Token.TabColor ?? Colors.LightGray)
        };

        public ICommand Activate { get; }
    }
}
