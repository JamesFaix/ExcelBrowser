using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using static ExcelBrowser.Interop.FetchStatus;

namespace ExcelBrowser.Interop {

    public enum FetchStatus { NotStarted, Fetching, Found, Error }

    public class Fetcher<T> : IDisposable {

        public Fetcher(
            Func<T> getValue,
            double timeoutSeconds = 5.0,
            double retrySeconds = 0.1,
            T defaultValue = default(T),
            Func<T, bool> valueFilter = null,
            Func<Exception, bool> exceptionFilter = null) {

            if (getValue == null) throw new ArgumentNullException(nameof(getValue));
            if (timeoutSeconds < 0) throw new ArgumentOutOfRangeException("Timeout seconds must be positive", nameof(timeoutSeconds));
            if (retrySeconds < 0) throw new ArgumentOutOfRangeException("Retry seconds must be positive", nameof(retrySeconds));
            Trace.WriteLine("Fetcher constructed");

            GetValue = getValue;
            TimeoutSeconds = timeoutSeconds;
            RetrySeconds = retrySeconds;
            DefaultValue = defaultValue;
            ValueFilter = valueFilter ?? (x => !Equals(x, this.DefaultValue));
            ExceptionFilter = exceptionFilter ?? (x => true);
            Status = NotStarted;
        }

        private readonly Func<T> GetValue;
        private readonly Func<T, bool> ValueFilter;
        private readonly Func<Exception, bool> ExceptionFilter;

        #region Properties
        public T DefaultValue { get; }

        public double TimeoutSeconds { get; }

        public double RetrySeconds { get; }

        public FetchStatus Status { get; private set; }

        private T result;
        public T Result {
            get {
                if (Status != Found)
                    throw new InvalidOperationException("Have not yet fetched.");
                else return this.result;
            }
        }

        public Exception Exception { get; private set; }

        private bool isDisposed;
        #endregion

        #region Events
        public event EventHandler Fetched;
        private void OnFetched() => Fetched?.Invoke(this, EventArgs.Empty);
        #endregion

        public void Fetch() {
            if (this.Status == Fetching)
                throw new InvalidOperationException("Already fetching.");

            this.Status = Fetching;
            new Task(FetchImpl).Start();
        }

        private void FetchImpl() {
            try {
                this.result = GetValueLoop();
                Status = Found;
            }
            catch (Exception x) {
                this.result = DefaultValue;
                Status = Error;
                Exception = x;
            }
            OnFetched();
        }

        private T GetValueLoop() {
            var sw = new Stopwatch();
            sw.Start();
            T result;
            while (sw.ElapsedMilliseconds / 1000 < TimeoutSeconds && !isDisposed) {
                if (TryGetValue(out result)) return result;
                else Wait();
            }
            return DefaultValue;
        }

        private bool TryGetValue(out T value) {
            try {
                Trace.WriteLine("Fetcher attempting to fetch value");

                value = GetValue();
                return ValueFilter(value);
            }
            catch (Exception x) {
                if (this.ExceptionFilter(x)) {
                    value = DefaultValue;
                    return false;
                }
                else throw x;
            }
        }

        private void Wait() => Thread.Sleep((int)(RetrySeconds * 1000));

        public override string ToString() {
            switch (Status) {
                case Found:
                    return $"{Status}: {Result}";
                case Error:
                    return $"{Status}: {Exception}";
                default:
                    return Status.ToString();
            }
        }

        public void Dispose() {
            this.isDisposed = true;
            this.result = DefaultValue;
        }
    }
}