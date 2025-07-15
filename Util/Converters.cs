using System.Globalization;
using System.Windows;
using System.Windows.Data;
using static System.Windows.Visibility;

namespace Util {
    public abstract class ReadOnlyConverterBase : IValueConverter {
        protected readonly object UnsetValue = DependencyProperty.UnsetValue;
        public abstract object Convert(object value, Type targetType, object parameter, CultureInfo culture);
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => UnsetValue;
    }
    public abstract class ReadOnlyMultiConverterBase : IMultiValueConverter {
        protected readonly object UnsetValue = DependencyProperty.UnsetValue;
        public abstract object Convert(object[] values, Type targetType, object parameter, CultureInfo culture);
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) => [UnsetValue];
    }

    public class TruthyBooleanConverter: ReadOnlyConverterBase {
        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
            value switch {
                string s => !s.IsNullOrWhitespace(),
                Uri => true,
                bool b => b,
                null => false,
                _ when value.GetType().UnderlyingIfNullable().IsNumeric() => ((dynamic)value) != 0,
                _ => throw new NotImplementedException()
            };
    }
}
