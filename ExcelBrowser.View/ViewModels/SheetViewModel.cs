using System.Collections.ObjectModel;
using System.Drawing;
using System;

namespace ExcelBrowser.ViewModels {

    public class SheetViewModel {

        public SheetViewModel() {
            Windows = new ObservableCollection<SheetWindowViewModel>();
        }

        public string Name { get; set; }
        
        public Color TabColor { get; set; }

        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public ObservableCollection<SheetWindowViewModel> Windows { get; set; }
    }
}
