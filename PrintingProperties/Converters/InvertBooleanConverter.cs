using System.Globalization;
using System.Windows.Data;
using System;

namespace PrintingProperties.Converters;

public class InvertBooleanConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue)
        {
            return !boolValue;
        }

        // Default handling if the value is not a bool?
        return Binding.DoNothing;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue)
        {
            return !boolValue;
        }

        // Default handling if the value is not a bool?
        return Binding.DoNothing;
    }
}
