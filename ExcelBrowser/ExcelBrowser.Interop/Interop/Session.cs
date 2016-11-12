using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using xlApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelBrowser.Interop {

    public class Session {

        public static Session Current => new Session(Process.GetCurrentProcess().SessionId);

        public Session(int sessionId) {
            Debug.WriteLine("");
            Debug.WriteLine("Session.Constructor");

            SessionId = sessionId;
        }

        private IEnumerable<Process> Processes =>
            Process.GetProcessesByName("EXCEL")
            .Where(p => p.SessionId == this.SessionId);

        private static Fallible<xlApp> TryGetApp(Process process) =>
            new Fallible<xlApp>(() => process.AsExcelApp());

        public int SessionId { get; }

        public IEnumerable<int> ProcessIds =>
            Processes
            .Select(p => p.Id)
            .ToArray();

        public IEnumerable<int> UnreachableProcessIds =>
            ProcessIds
            .Except(Apps.Select(a => a.AsProcess().Id))
            .ToArray();

        public IEnumerable<xlApp> Apps =>
            Processes
            .Select(TryGetApp)
            .Values()
            .Where(a => a.AsProcess().IsVisible())
            .ToArray();

        public xlApp TopMost {
            get {
                int? topMostId = Apps.Select(p => p.AsProcess()).TopMost()?.Id;
                return topMostId.HasValue
                    ? Apps.Single(a => a.AsProcess().Id == topMostId)
                    : null;
            }
        }

        public xlApp Primary => AppFactory.PrimaryInstance;
    }
}
