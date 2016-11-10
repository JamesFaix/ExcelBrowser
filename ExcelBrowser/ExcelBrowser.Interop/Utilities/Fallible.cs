using System;

namespace ExcelBrowser {

    /// <summary>Represents either a value 
    /// OR an exception that was thrown when attempting to acquire the value.</summary>
    /// <typeparam name="T">Type of value.</typeparam>
    public class Fallible<T> : IEquatable<Fallible<T>> {

        /// <summary>Initializes a default instance, with <c>default(T)</c> as a value.</summary>
        public Fallible() {
            //value defaults to default(T)
            //exception defaults to null
        }

        /// <summary>Initializes a new instance with the given value.</summary>
        /// <param name="value">The value.</param>
        public Fallible(T value) : this() {
            this.value = value;
        }

        /// <summary>Initializes a new instance with the given exception.</summary>
        /// <param name="error">The exception.</param>
        public Fallible(Exception error) : this() {
            Requires.NotNull(error, nameof(error));
            Error = error;
        }

        public Fallible(Func<T> getValue) : this() {
            Requires.NotNull(getValue, nameof(getValue));
            try {
                this.value = getValue();
            }
            catch(Exception x) {
                Error = x;
            }
        }

        /// <summary>Gets the value.</summary>
        /// <exception cref="InvalidOperationException">Cannot get Value if HasValue is false.</exception>
        public T Value {
            get {
                Requires.Rule(HasValue, "Cannot get Value if HasValue is false.");
                return value;
            }
        }
        private T value;

        /// <summary>Gets the exception.</summary>
        public Exception Error { get; }

        /// <summary>Gets a value indicating whether this instance has value.</summary>
        public bool HasValue => Error == null;

        #region Equality

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>
        /// true if the current object is equal to the <paramref name="other" /> parameter; otherwise, false.
        /// </returns>
        public bool Equals(Fallible<T> other) => (other != null)
            && Equals(Error, other.Error)
            && Equals(Value, other.Value);

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" />, is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
        public override bool Equals(object obj) => Equals(obj as Fallible<T>);

        /// <summary>
        /// Implements the operator ==.
        /// </summary>
        /// <param name="a">a.</param>
        /// <param name="b">The b.</param>
        /// <returns>
        /// The result of the operator.
        /// </returns>
        public static bool operator ==(Fallible<T> a, Fallible<T> b) {
            if (a == null) return b == null;
            return a.Equals(b);
        }

        /// <summary>
        /// Implements the operator !=.
        /// </summary>
        /// <param name="a">a.</param>
        /// <param name="b">The b.</param>
        /// <returns>
        /// The result of the operator.
        /// </returns>
        public static bool operator !=(Fallible<T> a, Fallible<T> b) {
            if (a == null) return b != null;
            return !a.Equals(b);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode() {
            if (Error != null) return Error.GetHashCode();
            if (Value != null) return Value.GetHashCode();
            return 0;
        }

        #endregion

        /// <summary>Returns a <see cref="string" /> that represents this instance.</summary>
        public override string ToString() =>
            (Error == null)
                ? $"Fallible{{{Value}}}"
                : $"Fallible{{{Error.GetType()}: {Error.Message}}}";
    }
}
