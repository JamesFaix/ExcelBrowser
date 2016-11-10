using System;

namespace ExcelBrowser.Monitoring {

    /// <summary>Represents the change of a variable from one value to another.</summary>
    /// <typeparam name="T">Type of variable.</typeparam>
    public class ValueChange<T> {

        /// <summary>Initializes a new instance.</summary>
        /// <param name="oldValue">The old value.</param>
        /// <param name="newValue">The new value.</param>
        public ValueChange(T oldValue, T newValue) {
            OldValue = oldValue;
            NewValue = newValue;
        }

        /// <summary>Gets the old value.</summary>
        public T OldValue { get; }

        /// <summary>Gets the new value.</summary>
        public T NewValue { get; }

        /// <summary>Returns a <see cref="string" /> that represents this instance.</summary>
        public override string ToString() => $"ValueChange: {{{OldValue}}} => {{{NewValue}}}";

        public bool IsDifferent(Func<T, T, bool> equalityComparison = null) {
            if (equalityComparison == null) {
                return !Equals(OldValue, NewValue);
            }
            else {
                return !equalityComparison(OldValue, NewValue);
            }
        }

        public ValueChange<TResult> Select<TResult>(Func<T, TResult> selector) {
            Requires.NotNull(selector, nameof(selector));
            return new ValueChange<TResult>(selector(OldValue), selector(NewValue));
        }
    }

    public static class ValueChange {

        public static ValueChange<T> Create<T>(T oldValue, T newValue) {
            return new ValueChange<T>(oldValue, newValue);
        }
    }
}
