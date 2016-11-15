using System.Collections.Generic;
using ExcelBrowser.Controller;
using ExcelBrowser.Model;
using System.ComponentModel;
using System.Diagnostics;

namespace ExcelBrowser.ViewModels {

    public class ViewModelUpdater : INotifyPropertyChanged {

        public ViewModelUpdater(SessionMonitor monitor) {
            Requires.NotNull(monitor, nameof(monitor));
            ViewModel = new SessionViewModel();

            this.monitor = monitor;
            monitor.SessionChanged += SessionChanged;
            UpdateViewModel();
        }

        private readonly SessionMonitor monitor;

        public event PropertyChangedEventHandler PropertyChanged;

        public SessionViewModel ViewModel { get; set; }

        private void UpdateViewModel() {
            ViewModel = ViewModelFactory.ConvertSession(monitor.Session);
        }

        private void SessionChanged(object sender, EventArgs<IEnumerable<Change>> e) {
            UpdateViewModel();
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("ViewModel"));
            Debug.WriteLine("ViewModel updated");
        }        
        
        //private void SessionChanged(object sender, EventArgs<IEnumerable<Change>> e) {
        //    foreach (var change in e.Value) {
        //        var idType = change.Id.GetType();
        //        var type = change.Type;

        //        if (idType == typeof(SessionId)) {
        //            switch (type) {
        //                case ChangeType.SessionStart:
        //                    ViewModel = ViewModelFactory.ConvertSession(monitor.Session);
        //                    break;

        //                case ChangeType.Add:
        //                    ViewModel.Apps.Add(change.)
        //            }
        //        }
        //        else if (idType == typeof(AppId)) {
        //            switch (type) {

        //            }
        //        }
        //        else if(idType == typeof(BookId)) {

        //        }
        //        else if (idType == typeof(SheetId)) {

        //        }
        //        else if(idType == typeof(WindowId)) {

        //        }
        //        else {
        //            throw new NotSupportedException("Invalid ID type");
        //        }
        //    }
        //}
    }
}
