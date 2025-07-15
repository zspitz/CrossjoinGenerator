namespace Util; 
public static class IEnumerableTExtensions {
    public static IEnumerable<T> ForEach<T>(this IEnumerable<T> src, Action<T> action) {
        foreach (var item in src) {
            action(item);
        }
        return src;
    }
    public static void AddRangeTo<T>(this IEnumerable<T> src, ICollection<T> dest) => dest.AddRange(src);
    public static string Joined<T>(this IEnumerable<T> source, string delimiter = ",", Func<T, string>? selector = null) =>
        source is null ? "" :
        selector is null ? string.Join(delimiter, source) :
        string.Join(delimiter, source.Select(selector));
}
