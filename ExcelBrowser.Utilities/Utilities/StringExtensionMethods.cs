using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelBrowser {

    public static class StringExtensionMethods {

        public static string ReplaceChars(this string source, string chars, string replacement = "") {
            Requires.NotNull(source, nameof(source));
            Requires.NotNull(chars, nameof(chars));
            Requires.NotNull(replacement, nameof(replacement));

            var sb = new StringBuilder();
            foreach (var c in source) {
                if (chars.Contains(c)) {
                    sb.Append(replacement);
                }
                else {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        public static string Replace(this string source, Regex pattern, string replacement = "") {
            Requires.NotNull(source, nameof(source));
            Requires.NotNull(pattern, nameof(pattern));
            Requires.NotNull(replacement, nameof(replacement));

            return pattern.Replace(source, replacement);
        }
    }
}
