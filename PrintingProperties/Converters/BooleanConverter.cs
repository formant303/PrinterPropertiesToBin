using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows;

namespace PrintingProperties.Converters;

public class BooleanConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue)
        {
            // Customize the conversion logic based on your needs
            return boolValue ? Visibility.Visible : Visibility.Collapsed;
        }

        // Return a default value if the conversion fails
        return DependencyProperty.UnsetValue;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        // Optionally implement ConvertBack if you need two-way binding
        throw new NotImplementedException();
    }
}
