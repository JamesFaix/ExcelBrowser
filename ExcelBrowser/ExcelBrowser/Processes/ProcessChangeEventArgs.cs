using System;

namespace ExcelBrowser.Processes {

    public class ProcessChangeEventArgs : EventArgs {

        public ProcessChangeEventArgs(int[] startedProcessIds, int[] stoppedProcessIds) : base() {
            StartedProcessIds = startedProcessIds ?? new int[0];
            StoppedProcessIds = stoppedProcessIds ?? new int[0];
        }

        public int[] StartedProcessIds { get; }
        public int[] StoppedProcessIds { get; }
    }
}
