using System.Globalization;
using Util;
using static System.Windows.Media.Brushes;
using static CrossjoinGenerator.ProcessState;

namespace CrossjoinGenerator;
public class ProcessStateToBrushConverter : ReadOnlyConverterBase {
    public override object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value switch {
            Error => Red,
            Warning => Orange,
            Success => Green,
            _ => parameter ?? UnsetValue,
        };
}

