using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ExcelBrowser.Interop {

    /// <summary>
    /// Provides methods to get the Z coordinate the main windows of processes.
    /// </summary>
    public static class ProcessWindowUtil {
        
        /// <summary>
        /// Gets the "Z" value of the main window. Lower Z's are closer to the screen.
        /// </summary>
        /// <param name="process">The process.</param>
        /// <returns>"Z" value of process's main window.</returns>
        public static int MainWindowZ(this Process process) {
            Requires.NotNull(process, nameof(process));            
            return WindowHandleUtil.GetWindowZ(process.MainWindowHandle);
        }

        /// <summary>
        /// Orders the sequence of processes by the "Z" value of their main window.
        /// </summary>
        /// <param name="processes">The processes.</param>
        /// <returns>Sequence of processes, ordered by the "Z" value of their main window.</returns>
        public static IEnumerable<Process> OrderByZ(this IEnumerable<Process> processes) {
            Requires.NotNull(processes, nameof(processes));

            return processes
                .Select(p => new {
                    Process = p,
                    Z = MainWindowZ(p)
                })
                .Where(x => x.Z > 0) //Filter hidden instances
                .OrderBy(x => x.Z) //Sort by z value
                .Select(x => x.Process);
        }

        /// <summary>
        /// Returns the process with the top-most main window.
        /// </summary>
        /// <param name="processes">The processes.</param>
        /// <returns>Process with top-most main window.</returns>
        public static Process TopMost(this IEnumerable<Process> processes) {
            Requires.NotNull(processes, nameof(processes));
            return OrderByZ(processes).FirstOrDefault();
        }
    }
}
