namespace ExcelBrowser.ViewModels {

    public class BookPanelViewModel {

        public string Name { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public bool IsAddIn { get; set; }
        public int WindowCount { get; set; }
        public SheetPanelViewModel[] Sheets { get; set; }
        public WindowsPanelViewModel Windows { get; set; }
    }
}
