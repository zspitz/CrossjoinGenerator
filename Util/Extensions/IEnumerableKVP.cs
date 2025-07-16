using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Util;

public static class IEnumerableKVPExtensions {
    public static IEnumerable<TResult> SelectKVP<TKey, TValue, TResult>(this IEnumerable<KeyValuePair<TKey, TValue>> src, Func<TKey, TValue, TResult> selector) =>
        src.Select(kvp => selector(kvp.Key, kvp.Value));
}
