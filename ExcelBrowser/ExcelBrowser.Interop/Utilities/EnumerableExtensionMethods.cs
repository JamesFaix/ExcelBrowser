using System.Collections.Generic;
using System.Linq;

namespace ExcelBrowser {
    public static class EnumerableExtensionMethods {

        public static int IndexOf<T>(this IEnumerable<T> sequence, T value) {
            Requires.NotNull(sequence, nameof(sequence));

            var list = sequence as IList<T>;
            if (list != null) { //Optimize for ILists
                return list.IndexOf(value);
            }
            else {
                var index = 0;
                foreach (var item in sequence) {
                    if (Equals(value, item)) return index;
                    index++;
                }
                return -1; //Default to -1
            }
        }

        public static IEnumerable<T> ConcatIf<T>(this IEnumerable<T> first, bool condition, IEnumerable<T> second) {
            Requires.NotNull(first, nameof(first));
            Requires.NotNull(second, nameof(second));
            return condition ? first.Concat(second) : first;
        }

        public static IEnumerable<T> ConcatIf<T>(this IEnumerable<T> sequence, bool condition, T element) {
            Requires.NotNull(sequence, nameof(sequence));
            return condition ? sequence.Concat(new[] { element }) : sequence;
        }
    }
}
