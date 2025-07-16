namespace Util; 
public static class IEnumerableTExtensions {
    public static IEnumerable<T> ForEach<T>(this IEnumerable<T> src, Action<T> action) {
        foreach (var item in src) {
            action(item);
        }
        return src;
    }
    public static IEnumerable<T> ForEach<T>(this IEnumerable<T> src, Action<T, int> action) {
        var current = 0;
        foreach (var item in src) {
            action(item, current);
            current += 1;
        }
        return src;
    }
    public static void AddRangeTo<T>(this IEnumerable<T> src, ICollection<T> dest) => dest.AddRange(src);
    public static string Joined<T>(this IEnumerable<T> source, string delimiter = ",", Func<T, string>? selector = null) =>
        source is null ? "" :
        selector is null ? string.Join(delimiter, source) :
        string.Join(delimiter, source.Select(selector));
    public static IEnumerable<T> SelectMany<T>(this IEnumerable<IEnumerable<T>> src) => src.SelectMany(x => x);
    public static IEnumerable<TResult> SelectT<T1, T2, TResult>(this IEnumerable<(T1, T2)> src, Func<T1, T2, TResult> selector) =>
        src.Select(x => selector(x.Item1, x.Item2));

}
