using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// True =&gt; Collapsed, false =&gt; Visible. The inverse of WPF's built-in
    /// BooleanToVisibilityConverter, for showing a control only when a flag is
    /// off (e.g. the inline path line when a row is not expanded). v2.4 C1.
    /// </summary>
    [ValueConversion(typeof(bool), typeof(Visibility))]
    public sealed class InverseBoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && b ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
