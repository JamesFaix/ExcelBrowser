using System;
using System.Windows.Input;

namespace ExcelBrowser.ViewModels {
    /// <summary>
    /// A command whose sole purpose is to 
    /// relay its functionality to other
    /// objects by invoking delegates. The
    /// default return value for the CanExecute
    /// method is 'true'.
    /// </summary>
    public class RelayCommand : ICommand {

        /// <summary>
        /// Creates a new command.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        /// <param name="canExecute">The execution status logic.</param>
        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null) {
            Requires.NotNull(execute, nameof(execute));

            this.execute = execute;
            this.canExecute = canExecute ?? AlwaysTrue;
        }

        private static bool AlwaysTrue(object obj) => true;

        public bool CanExecute(object parameters) => canExecute(parameters);
        private readonly Func<object, bool> canExecute;

        public void Execute(object parameters) => execute(parameters);
        private readonly Action<object> execute;

        public event EventHandler CanExecuteChanged {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}