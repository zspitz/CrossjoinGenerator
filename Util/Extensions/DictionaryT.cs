namespace Util;
public static class DictionaryTExtensions {
    public static IEnumerable<(TKey, TValue)> Flatten<TKey, TValue,TCollection>(this IDictionary<TKey, TCollection> dict) 
            where TKey : notnull
            where TCollection : IEnumerable<TValue> 
        => 
            dict.SelectMany(kvp => kvp.Value.Select(v => (kvp.Key, v)));
}
