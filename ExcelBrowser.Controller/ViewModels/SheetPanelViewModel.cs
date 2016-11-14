using System.Drawing;

namespace ExcelBrowser.ViewModels {

    public class SheetPanelViewModel {

        public string Name { get; set; }
        public Color TabColor { get; set; }
        public int WindowCount { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public bool[] IsActiveInWindow { get; set; }

    }
}
