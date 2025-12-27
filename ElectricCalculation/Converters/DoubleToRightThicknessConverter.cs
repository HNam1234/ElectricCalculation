using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ElectricCalculation.Converters
{
    public sealed class DoubleToRightThicknessConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var number = value switch
            {
                double d => d,
                float f => f,
                decimal m => (double)m,
                int i => i,
                long l => l,
                _ => 0d
            };

            if (double.IsNaN(number) || double.IsInfinity(number) || number < 0)
            {
                number = 0;
            }

            return new Thickness(0, 0, number, 0);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}

