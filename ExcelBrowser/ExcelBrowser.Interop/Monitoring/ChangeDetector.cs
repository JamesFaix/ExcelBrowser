using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace ExcelBrowser.Monitoring {

    //Does not need to implement Dispose(bool disposing) because derived classes can just call Dispose()
    [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]

    public class ChangeDetector<T> : IDisposable {

        public ChangeDetector(Func<T> getValue, double refreshSeconds, 
            Func<T, T, bool> valueEqualityComparison = null)
            : base() {
            Requires.NotNull(getValue, nameof(getValue));
            Requires.Positive(refreshSeconds, nameof(refreshSeconds));
            Debug.WriteLine("ChangeDetector.Constructor");

            RefreshSeconds = refreshSeconds;
            this.valueEqualityComparison = valueEqualityComparison ?? ((x, y) => Equals(x, y));

            this.valueChecker = new ValueChecker<T>(
                getValue: getValue,
                refreshSeconds: refreshSeconds);

            valueChecker.Next += ValueReceived;
        }

        private readonly ValueChecker<T> valueChecker;
        private readonly Func<T, T, bool> valueEqualityComparison;

        private T previousValue;

        public double RefreshSeconds { get; }

        private void ValueReceived(object sender, EventArgs<Fallible<T>> e) {
            if (e.Value.HasValue) { //Check if Fallible has actual value 
                var currentValue = e.Value.Value;

                if (!valueEqualityComparison(currentValue, this.previousValue)) {
                    var oldPrevious = previousValue;
                    previousValue = currentValue;
                    Debug.WriteLine("ChangeDetector: Change detected");
                    OnChanged(new ValueChange<T>(oldPrevious, currentValue));
                }
                else {
                    Debug.Print("ChangeDetector: No change");
                }
            }
            else {
                Debug.WriteLine("ChangeDetector: " + e.Value.Error.Message.Replace("\n", " "));
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]
        public void Dispose() {
            valueChecker.Dispose();
        }

        public event EventHandler<EventArgs<ValueChange<T>>> Changed;
        private void OnChanged(ValueChange<T> change) => Changed?.Invoke(this, new EventArgs<ValueChange<T>>(change));
    }
}
