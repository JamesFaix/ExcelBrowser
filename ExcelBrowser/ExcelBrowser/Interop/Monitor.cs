using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelBrowser.Interop {

    public class Monitor<T> {

        public Monitor(
            Func<T> getValue,
            double retrySeconds = 0.1,
            TimeSpan? timeout = null,
            Func<T, T, bool> valueComparison = null,
            T initialValue = default(T)) {

            if (getValue == null) throw new ArgumentNullException(nameof(getValue));
            if (retrySeconds < 0) throw new ArgumentOutOfRangeException("Retry seconds must be positive", nameof(retrySeconds));
            Trace.WriteLine("Creating monitor");

            GetValue = getValue;
            ValueComparison = valueComparison ?? EqualityComparer<T>.Default.Equals;
            RetrySeconds = retrySeconds;
            Timeout = timeout;
            Value = initialValue;

            new Task(GetValueLoop).Start();
        }

        private readonly Func<T> GetValue;
        private readonly Func<T, T, bool> ValueComparison;

        #region Properties

        public double RetrySeconds { get; }

        public TimeSpan? Timeout { get; }

        public T Value { get; set; }

        #endregion

        #region Events
        public event EventHandler<ValueChangedEventArgs<T>> ValueChanged;
        private void OnValueChanged(T oldValue, T newValue) =>
            ValueChanged?.Invoke(this, new ValueChangedEventArgs<T>(oldValue, newValue));
        #endregion

        private void GetValueLoop() {
            var sw = new Stopwatch();
            sw.Start();
            while (!IsPastTimeout(sw.ElapsedMilliseconds)) {
                CheckAndUpdate();
                Wait();
            }
        }

        private bool IsPastTimeout(long milliseconds) =>
            Timeout.HasValue &&
            milliseconds > Timeout.Value.TotalMilliseconds;

        private void CheckAndUpdate() {
            Trace.WriteLine("Monitor checking for update");

            var newValue = GetValue();
            if (!ValueComparison(this.Value, newValue)) {
                OnValueChanged(this.Value, newValue);
                this.Value = newValue;
            }
        }

        private void Wait() => Thread.Sleep((int)(RetrySeconds * 1000));

        public override string ToString() => Value.ToString();
    }
}