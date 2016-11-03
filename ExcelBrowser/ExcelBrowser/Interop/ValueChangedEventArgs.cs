using System;

namespace ExcelBrowser.Interop {

    public class ValueChangedEventArgs<T> : EventArgs {
        public ValueChangedEventArgs(T oldValue, T newValue)
            : base() {
            OldValue = oldValue;
            NewValue = newValue;
        }

        public T OldValue { get; }
        public T NewValue { get; }
    }
}
