namespace ExcelBrowser.ViewModels {

    public class BookWindowViewModel {

        public int Index { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }

        public string Label => $"[{Index}]";
    }
}
