using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBrowser {

    /// <summary>Provides methods for common precondition and argument validation assertions.</summary>
    public static class Requires {

        /// <summary>Throws an <see cref="ArgumentNullException"/> if the given object is null.</summary>
        /// <typeparam name="T">Type of object.</typeparam>
        /// <param name="obj">The object.</param>
        /// <param name="name">The object name.</param>
        /// <exception cref="ArgumentNullException">Thrown if object is null.</exception>
        public static void NotNull<T>(T obj, string name) {
            if (Equals(obj, null))
                throw new ArgumentNullException(name);
        }

        public static void NotContainsNull<T>(IEnumerable<T> sequence, string name) where T: class {
            if (sequence.Contains(null))
                throw new ArgumentException("Sequence cannot contain null.", name);
        }

        /// <summary>Throws an <see cref="ArgumentException"/> if the given object is the default value for the given type.</summary>
        /// <typeparam name="T">Type of object.</typeparam>
        /// <param name="obj">The object.</param>
        /// <param name="name">The object name.</param>
        /// <exception cref="ArgumentException">Thrown if object is the default value for the given type.</exception>
        public static void NotDefault<T>(T obj, string name, string message = null) {
            if (Equals(obj, default(T)))
                throw new ArgumentException(message ?? "Argument cannot have default value.", name);
        }

        /// <summary>Throws an <see cref="ArgumentOutOfRangeException"/> if the given value is less than the default value for the given type.</summary>
        /// <typeparam name="T">Type of value.</typeparam>
        /// <param name="obj">The value.</param>
        /// <param name="name">The value name.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if argument is less than the default value for the given type.</exception>
        public static void Positive<T>(T obj, string name) where T : struct, IComparable {
            if (obj.CompareTo(default(T)) < 1)
                throw new ArgumentOutOfRangeException("Argument must be positive.", name);
        }

        /// <summary>Throws an <see cref="ArgumentOutOfRangeException"/> if the given value is not null and is less than the default value for the given type.</summary>
        /// <typeparam name="T">Type of value.</typeparam>
        /// <param name="obj">The value.</param>
        /// <param name="name">The value name.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if argument is not null and is less than the default value for the given type.</exception>
        public static void Positive<T>(T? obj, string name) where T : struct, IComparable {
            if (obj.HasValue && obj.Value.CompareTo(default(T)) < 1)
                throw new ArgumentOutOfRangeException("Argument must be positive.", name);
        }

        /// <summary>Throws an <see cref="InvalidOperationException"/> if the given condition is false.</summary>
        /// <param name="condition">Condition to test.</param>
        /// <param name="message">Message describing rule that was broken.</param>
        /// <exception cref="InvalidOperationException">Thrown if condition is false.</exception>
        public static void Rule(bool condition, string message) {
            if (!condition) throw new InvalidOperationException(message);
        }

        public static Exception ValidEnum(int value) => new InvalidOperationException("Invalid enum value. (" + value + ")");
    }
}
