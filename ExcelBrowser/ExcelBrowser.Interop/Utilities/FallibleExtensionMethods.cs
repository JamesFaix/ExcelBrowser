using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBrowser {

    public static class FallibleExtensionMethods {

        public static IEnumerable<T> Values<T>(this IEnumerable<Fallible<T>> sequence) {
            Requires.NotNull(sequence, nameof(sequence));
            return sequence.Where(f => f.HasValue).Select(f => f.Value);
        }

        public static IEnumerable<Exception> Errors<T>(this IEnumerable<Fallible<T>> sequence) {
            Requires.NotNull(sequence, nameof(sequence));
            return sequence.Where(f => !f.HasValue).Select(f => f.Error);
        }

    }
}
