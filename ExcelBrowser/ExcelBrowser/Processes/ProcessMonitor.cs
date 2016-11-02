using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelBrowser.Processes {

    public class ProcessMonitor {

        /// <summary>Initializes a new instance of the <see cref="ProcessUtilityBase"/> class. </summary>
        /// <param name="sessionId">ID of session to monitor. Defaults to current session.</param>
        public ProcessMonitor(int? sessionId = null) {
            this.SessionId = sessionId ?? Process.GetCurrentProcess().SessionId;
            this.refreshMilliseconds = 500;
            Task.Run((Action)CheckForProcessChanges);
        }

        /// <summary>Gets the session identifier. </summary>
        public int SessionId { get; private set; }

        /// <summary>Gets or sets the number of milliseconds between checking for process changes. </summary>
        /// <exception cref="ArgumentOutOfRangeException">RefreshMilliseconds must be greater than 0.</exception>
        public int RefreshMilliseconds {
            get { return refreshMilliseconds; }
            set {
                if (value <= 0) throw new ArgumentOutOfRangeException("RefreshMilliseconds must be greater than 0.");
                refreshMilliseconds = value;
            }
        }
        /// <summary>The number of milliseconds between checking for process changes. </summary>
        private int refreshMilliseconds;

        public int[] ProcessIds { get; private set; } = new int[0];

        public Func<string, bool> NameFilter {
            get { return nameFilter; }
            set { nameFilter = value ?? defaultNameFilter; }
        }
        private Func<string, bool> nameFilter = defaultNameFilter;
        private static Func<string, bool> defaultNameFilter = (x => true);

        private int[] GetCurrentProcessIds() {
            return Process.GetProcesses()
                .Where(p => nameFilter(p.ProcessName))
                .Where(p => p.SessionId == this.SessionId)
                .Select(p => p.Id)
                .ToArray();
        }

        /// <summary>Checks for process changes, then sleeps the thread. </summary>
        private void CheckForProcessChanges() {
            while (true) {
                var current = GetCurrentProcessIds();
                var previous = this.ProcessIds;

                //Compare current and previous
                var started = current.Except(previous).OrderBy(x => x).ToArray();
                var stopped = previous.Except(current).OrderBy(x => x).ToArray();

                //Update internal state before publishing event, so callbacks can access clean data.
                this.ProcessIds = current;

                //Publish any changes
                if (started.Any() || stopped.Any()) OnProcessChange(started, stopped);

                //Wait
                Thread.Sleep(refreshMilliseconds);
            }
        }

        public event EventHandler<ProcessChangeEventArgs> ProcessChange;
        private void OnProcessChange(int[] startedProcessIds, int[] stoppedProcessIds) =>
            ProcessChange?.Invoke(this, new ProcessChangeEventArgs(startedProcessIds, stoppedProcessIds));
    }
}
