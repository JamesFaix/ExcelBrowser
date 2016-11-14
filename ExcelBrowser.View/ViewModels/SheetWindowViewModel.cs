using System.Windows.Media;

namespace ExcelBrowser.ViewModels {

    public class SheetWindowViewModel {
        
        public Color? TabColor { get; set; }
        public bool IsActive { get; set; }

        public Brush Background => new SolidColorBrush { Color = TabColor ?? Colors.LightGray };
    }
}
