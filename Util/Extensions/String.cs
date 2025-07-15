using System.Diagnostics.CodeAnalysis;

namespace Util; 

public static class StringExtensions {
    public static bool IsNullOrWhitespace([NotNullWhen(false)] this string? s) => string.IsNullOrWhiteSpace(s);
}
