using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelBrowser.Monitoring {

    /// <summary>
    /// Publishes the result of a given function at a given interval.
    /// </summary>
    /// <typeparam name="T">Type of value.</typeparam>
    /// <seealso cref="IDisposable" />
    public class ValueChecker<T> : IDisposable {

        public ValueChecker(Func<T> getValue, double refreshSeconds = 0.1, 
            Func<T, bool> valueFilter = null, 
            Func<Exception, bool> errorFilter = null) : base() {

            Requires.NotNull(getValue, nameof(getValue));
            Requires.Positive(refreshSeconds, nameof(refreshSeconds));

            this.getValue = getValue;
            this.valueFilter = valueFilter ?? (x => true);
            this.errorFilter = this.errorFilter ?? (x => true);
            this.RefreshSeconds = refreshSeconds;

            this.tokenSource = new CancellationTokenSource();
            var token = tokenSource.Token;

            this.task = Task.Run(() => CheckLoop(token));
        }

        private readonly Func<T> getValue;
        private readonly Func<T, bool> valueFilter;
        private readonly Func<Exception, bool> errorFilter;

        private bool isDisposed;
        private readonly CancellationTokenSource tokenSource;
        private readonly Task task;

        public double RefreshSeconds { get; }
        
        public void Dispose() {
            isDisposed = true;
            tokenSource.Cancel();
            Next = null;
        }

        private void CheckLoop(CancellationToken token) {            
            var ms = (int)(RefreshSeconds * 1000);

            while (!isDisposed) {
                token.ThrowIfCancellationRequested();
                
                try {
                    var value = getValue();
                    if (valueFilter(value)) OnNext(new Fallible<T>(value));
                }
                catch (Exception error) {
                    if (errorFilter(error)) OnNext(new Fallible<T>(error));
                }

                Thread.Sleep(ms);
            }
        }

        public event EventHandler<EventArgs<Fallible<T>>> Next;
        private void OnNext(Fallible<T> value) => Next?.Invoke(this, new EventArgs<Fallible<T>>(value));
    }
}
