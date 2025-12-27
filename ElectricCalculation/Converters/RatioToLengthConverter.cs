using System;
using System.Globalization;
using System.Windows.Data;

namespace ElectricCalculation.Converters
{
    public sealed class RatioToLengthConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values is not { Length: >= 2 })
            {
                return 0d;
            }

            var ratio = TryToDouble(values[0]);
            var maxLength = TryToDouble(values[1]);

            if (double.IsNaN(ratio) || double.IsInfinity(ratio))
            {
                ratio = 0;
            }

            if (double.IsNaN(maxLength) || double.IsInfinity(maxLength) || maxLength <= 0)
            {
                return 0d;
            }

            if (ratio < 0)
            {
                ratio = 0;
            }

            if (ratio > 1)
            {
                ratio = 1;
            }

            return ratio * maxLength;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        private static double TryToDouble(object value)
        {
            return value switch
            {
                double d => d,
                float f => f,
                decimal m => (double)m,
                int i => i,
                long l => l,
                string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) => parsed,
                _ => 0d
            };
        }
    }
}

