using System;
using System.Windows.Data;

namespace Telegraph.Converters
{
    public class UrgencyConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            int param = int.Parse(parameter.ToString());
            bool res = (int)value == param ? true : false;
            return res;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (bool)value ? parameter : null;
        }
    }
}
