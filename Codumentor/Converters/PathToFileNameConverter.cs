using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace Codumentor.Converters
{
    public class PathToFileNameConverter : MarkupExtension, IValueConverter
    {
        private static PathToFileNameConverter _instance;

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            if (_instance == null)
            {
                _instance = new PathToFileNameConverter();
            }
            return _instance;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is string path ? System.IO.Path.GetFileName(path) : value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
