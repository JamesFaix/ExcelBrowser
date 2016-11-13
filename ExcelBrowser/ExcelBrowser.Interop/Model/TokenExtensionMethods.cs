using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace ExcelBrowser.Model {

    internal static class TokenExtensionMethods {

        public static IEnumerable<TId> Ids<TId>(this IEnumerable<Token<TId>> tokens) =>
          tokens.Select(t => t.Id);

        public static IEnumerable<TToken> ExceptIds<TId, TToken>(this IEnumerable<TToken> tokens, IEnumerable<TId> ids)
            where TToken : Token<TId> =>
            tokens.Where(t => !ids.Contains(t.Id));

        public static IEnumerable<TToken> IntersectIds<TId, TToken>(this IEnumerable<TToken> tokens, IEnumerable<TId> ids)
           where TToken : Token<TId> =>
           tokens.Where(t => ids.Contains(t.Id));
    }
}
