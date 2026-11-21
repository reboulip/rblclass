using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Maps a bool to one of two <see cref="GridLength"/>s so a Grid's row and
    /// column sizes can follow the pane's horizontal/vertical layout mode
    /// (v2.13.0.0 A2): definitions cannot take triggers, but they can bind.
    /// </summary>
    public sealed class BoolToGridLengthConverter : IValueConverter
    {
        public GridLength TrueValue { get; set; } = GridLength.Auto;
        public GridLength FalseValue { get; set; } = GridLength.Auto;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && b ? TrueValue : FalseValue;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
