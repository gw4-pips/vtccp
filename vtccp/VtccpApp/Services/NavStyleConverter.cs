namespace VtccpApp.Services;

using System.Globalization;
using System.Windows;
using System.Windows.Data;

/// <summary>
/// Returns the active nav-button style when the bound CurrentPageKey matches the parameter.
/// Usage: Style="{Binding CurrentPageKey, Converter={StaticResource NavStyleConverter}, ConverterParameter=Session}"
/// </summary>
public sealed class NavStyleConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        string? current   = value as string;
        string? key       = parameter as string;
        bool    isActive  = string.Equals(current, key, StringComparison.OrdinalIgnoreCase);

        return isActive
            ? Application.Current.FindResource("NavButtonActive")
            : Application.Current.FindResource("NavButton");
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => Binding.DoNothing;
}
