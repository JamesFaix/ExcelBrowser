using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelBrowser.Controller;
using ExcelBrowser;
using ExcelBrowser.Model;

namespace TestApp {
    class Program {
        static void Main(string[] args) {

            var updater = new SessionUpdater(refreshSeconds: 2);
            updater.Changed += SessionChanged;

            Console.Read();
        }

        private static void SessionChanged(object sender, EventArgs<IEnumerable<ModelChange>> e) {
            foreach (var change in e.Value) {
                Console.WriteLine(change);
            }
        }
    }
}
