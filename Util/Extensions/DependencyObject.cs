using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Util; 
public static class DependencyObjectExtensions {
    /// <summary>Sets the value of the <paramref name="property"/> only if it hasn't been explicitly set.</summary>
    public static bool SetIfDefault<T>(this DependencyObject o, DependencyProperty property, T value) {
        if (DependencyPropertyHelper.GetValueSource(o, property).BaseValueSource == BaseValueSource.Default) {
            o.SetValue(property, value);
            return true;
        }
        return false;
    }

}
