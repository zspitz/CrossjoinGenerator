using System.Numerics;

namespace Util; 

public static class TypeExtensions {
    public static Type UnderlyingIfNullable(this Type type) => Nullable.GetUnderlyingType(type) ?? type;

    private static readonly Dictionary<Type, bool> numericTypes = new() {
        [typeof(byte)] = true,
        [typeof(short)] = true,
        [typeof(int)] = true,
        [typeof(long)] = true,
        [typeof(sbyte)] = true,
        [typeof(ushort)] = true,
        [typeof(uint)] = true,
        [typeof(ulong)] = true,
        [typeof(BigInteger)] = true,
        [typeof(float)] = false,
        [typeof(double)] = false,
        [typeof(decimal)] = false
    };

    public static bool IsNumeric(this Type type) => numericTypes.ContainsKey(type);
}
