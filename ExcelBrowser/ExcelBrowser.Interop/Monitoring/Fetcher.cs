using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Diagnostics.CodeAnalysis;

namespace ExcelBrowser.Monitoring {
    
    //Does not need to implement Dispose(bool disposing) because derived classes can just call Dispose()
    [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]

    public class Fetcher<T> : IDisposable {

        public Fetcher(Func<T> getValue, double refreshSeconds = 0.1, double? timeoutSeconds = 1,
            Func<T, bool> valueFilter = null, Func<Exception, bool> errorFilter = null) {
            Requires.NotNull(getValue, nameof(getValue));
            Requires.Positive(refreshSeconds, nameof(refreshSeconds));
            Requires.Positive(timeoutSeconds, nameof(timeoutSeconds));

            RefreshSeconds = refreshSeconds;
            TimeoutSeconds = timeoutSeconds;
            Status = FetchStatus.NotStarted;

            this.createChecker = () => new ValueChecker<T>(
                getValue: getValue,
                refreshSeconds: refreshSeconds,
                valueFilter: valueFilter ?? (x => true),
                errorFilter: errorFilter ?? (x => true));

            this.task = new Task<Fallible<T>>(FetchImpl);
        }

        #region Properties, fields, delegates

        private readonly Task<Fallible<T>> task;
        private readonly Func<ValueChecker<T>> createChecker;

        public double RefreshSeconds { get; }
        public double? TimeoutSeconds { get; }

        public FetchStatus Status { get; private set; }

        public Fallible<T> Result {
            get {
                Requires.Rule(Status == FetchStatus.Complete, "Cannot get Result if Status is not Complete.");
                return result;
            }
        }
        private Fallible<T> result;

        private ValueChecker<T> valueChecker;

        #endregion

        #region Fetch

        public Fallible<T> Fetch() {
            Requires.Rule(Status == FetchStatus.NotStarted, "Fetcher can only fetch once.");

            this.Status = FetchStatus.Started;
            this.task.Wait();
            return task.Result;
        }

        public async Task<Fallible<T>> FetchAsync() {
            Requires.Rule(Status == FetchStatus.NotStarted, "Fetcher can only fetch once.");

            this.Status = FetchStatus.Started;
            return await this.task;
        }

        private Fallible<T> FetchImpl() {
            this.valueChecker = createChecker();
            valueChecker.Next += ValueFound;
            Loop();
            return Result;
        }

        private void Loop() {
            if (TimeoutSeconds.HasValue) {
                var sw = new Stopwatch();
                sw.Start();

                bool isTimedOut = false;
                while (Status == FetchStatus.Started && !isTimedOut) {
                    isTimedOut = (sw.ElapsedMilliseconds / 1000) >= TimeoutSeconds.Value;
                }
                if (Status != FetchStatus.Complete) Timeout();
            }
            else {
                while (Status == FetchStatus.Started) {
                    //Do nothing
                }
            }
        }

        private void ValueFound(object sender, EventArgs<Fallible<T>> e) => OnComplete(e.Value);
        private void Timeout() => OnComplete(new Fallible<T>(new TimeoutException("Fetch timed out.")));
        public void Cancel() => OnComplete(new Fallible<T>(new OperationCanceledException("Fetch canceled.")));
        
        public event EventHandler<EventArgs<Fallible<T>>> Complete;
        private void OnComplete(Fallible<T> value) {
            valueChecker.Dispose();
            Status = FetchStatus.Complete;
            result = value;
            Complete?.Invoke(this, new EventArgs<Fallible<T>>(value));
        }

        #endregion

        [SuppressMessage("Microsoft.Design", "CA1063:ImplementIDisposableCorrectly")]
        public void Dispose() {
            valueChecker.Dispose();
        }
    }

    public static class Fetcher {

        public static Fallible<T> Fetch<T>(
            Func<T> getValue, double refreshSeconds = 0.1, double? timeoutSeconds = 1,
            Func<T, bool> valueFilter = null, Func<Exception, bool> errorFilter = null) {

            using (var fetcher = new Fetcher<T>(getValue, refreshSeconds, timeoutSeconds, valueFilter, errorFilter)) {
                return fetcher.Fetch();
            }
        }

        public static async Task<Fallible<T>> FetchAsync<T>(
            Func<T> getValue, double refreshSeconds = 0.1, double? timeoutSeconds = 1,
            Func<T, bool> valueFilter = null, Func<Exception, bool> errorFilter = null) {

            using (var fetcher = new Fetcher<T>(getValue, refreshSeconds, timeoutSeconds, valueFilter, errorFilter)) {
                return await fetcher.FetchAsync();
            }
        }
    }
}
