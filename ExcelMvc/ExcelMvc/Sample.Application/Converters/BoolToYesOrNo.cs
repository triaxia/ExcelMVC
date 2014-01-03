using System;
using System.Globalization;
using System.Windows.Data;

namespace Sample.Application.Converters
{
    public class BoolToYesOrNo : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null && System.Convert.ToBoolean(value) ? "Yes" : "No";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var svalue = value == null ? "" : System.Convert.ToString(value);
            return svalue.StartsWith("Y") || svalue.StartsWith("y");
        }
    }
}
