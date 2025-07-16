using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Util;
public static class ICollectionTExtensions {
    public static void AddRange<T>(this ICollection<T> dest, IEnumerable<T> toAdd) => toAdd.ForEach(x => dest.Add(x));
    public static object[,] To2DRow(this ICollection<string> source) {
        var to2DRow = new object[1, source.Count];
        int i = 0;
        foreach (var item in source) {
            to2DRow[0, i] = item;
        }
        return to2DRow;
    }
}
