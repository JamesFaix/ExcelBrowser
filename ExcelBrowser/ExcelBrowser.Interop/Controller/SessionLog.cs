using System;
using System.Collections.Generic;
using System.ComponentModel;
using ExcelBrowser.Model;
using System.Text.RegularExpressions;

namespace ExcelBrowser.Controller {

    public class SessionLog : INotifyPropertyChanged {

        public SessionLog(SessionMonitor monitor) {
            Requires.NotNull(monitor, nameof(monitor));

            this.monitor = monitor;
            monitor.SessionChanged += LogSessionChange;

            Text = "";
        }

        private readonly SessionMonitor monitor;

        public string Text { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void LogSessionChange(object sender, EventArgs<IEnumerable<ModelChange>> e) {
            foreach (var change in e.Value) {                
                Text += Format(change);
            }
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Text"));
        }

        private const string dateFormat = "hh:mm:ss.fff";
        private static Regex largeSpaces = new Regex(@"[\r\t\n]", RegexOptions.Compiled);
        private static Regex multipleSpaces = new Regex(@"\s{2,}", RegexOptions.Compiled);

        private string Format(ModelChange change) =>
            DateTime.Now.ToString(dateFormat) + " " +
            change.ToString()
            .Replace(largeSpaces)
            .Replace(multipleSpaces, " ")
            + Environment.NewLine;
    }
}
