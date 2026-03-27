namespace VtccpApp.Services;

using System.Globalization;
using System.Windows.Data;

/// <summary>Inverts a boolean value. Typically used to disable controls while IsRunning is true.</summary>
[ValueConversion(typeof(bool), typeof(bool))]
public sealed class InvertBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool b && !b;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool b && !b;
}
