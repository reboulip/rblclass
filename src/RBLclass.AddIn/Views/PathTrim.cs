using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RBLclass.Core;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Attached behavior that fills a <see cref="TextBlock"/> with a folder path
    /// truncated to keep its END visible (leading ellipsis) at the current
    /// control width, and sets the full path as the tooltip. WPF has no native
    /// leading/path ellipsis, so the text is measured and truncated through the
    /// unit-tested <see cref="PathEllipsis"/>; it is recomputed whenever the
    /// width changes.
    /// <para>
    /// The host TextBlock must be width-constrained for this to do anything -
    /// i.e. it should stretch inside a parent whose width is bounded (the owning
    /// ListBox disables horizontal scrolling), so <c>ActualWidth</c> reflects
    /// the available width rather than the full text width.
    /// </para>
    /// </summary>
    public static class PathTrim
    {
        public static readonly DependencyProperty FullPathProperty =
            DependencyProperty.RegisterAttached(
                "FullPath", typeof(string), typeof(PathTrim),
                new PropertyMetadata(null, OnFullPathChanged));

        public static void SetFullPath(DependencyObject d, string value) => d.SetValue(FullPathProperty, value);
        public static string GetFullPath(DependencyObject d) => (string)d.GetValue(FullPathProperty);

        /// <summary>
        /// Extra width (px) to keep clear on the right - e.g. for per-row buttons
        /// that sit beside the path. Subtracted from the available width so the
        /// truncated text never runs under them.
        /// </summary>
        public static readonly DependencyProperty RightReserveProperty =
            DependencyProperty.RegisterAttached(
                "RightReserve", typeof(double), typeof(PathTrim),
                new PropertyMetadata(0.0, OnRightReserveChanged));

        public static void SetRightReserve(DependencyObject d, double value) => d.SetValue(RightReserveProperty, value);
        public static double GetRightReserve(DependencyObject d) => (double)d.GetValue(RightReserveProperty);

        private static void OnRightReserveChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBlock tb) Apply(tb);
        }

        private static void OnFullPathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (!(d is TextBlock tb)) return;
            tb.SizeChanged -= OnSizeChanged; // idempotent re-hook (item containers are recycled)
            tb.SizeChanged += OnSizeChanged;
            Apply(tb);
        }

        private static void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (e.WidthChanged && sender is TextBlock tb) Apply(tb);
        }

        private static void Apply(TextBlock tb)
        {
            string full = GetFullPath(tb) ?? string.Empty;
            tb.ToolTip = string.IsNullOrEmpty(full) ? null : full;

            try
            {
                double available = AvailableWidth(tb);
                if (available <= 0)
                {
                    // Not laid out yet; show the full text and wait for SizeChanged.
                    tb.Text = full;
                    return;
                }

                var typeface = new Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch);
                double pixelsPerDip = VisualTreeHelper.GetDpi(tb).PixelsPerDip;

                tb.Text = PathEllipsis.TrimStart(full, available, candidate =>
                    new FormattedText(candidate, CultureInfo.CurrentUICulture, FlowDirection.LeftToRight,
                        typeface, tb.FontSize, Brushes.Black, pixelsPerDip).WidthIncludingTrailingWhitespace);
            }
            catch
            {
                // Never let a measuring hiccup blank the row or escape into WPF.
                tb.Text = full;
            }
        }

        /// <summary>
        /// Width available to the path text: from the TextBlock's left edge to
        /// the owning list's right content edge. Derived from the list (which is
        /// constrained to the pane) rather than the TextBlock's own ActualWidth,
        /// so it works regardless of whether the row stretches the TextBlock and
        /// it automatically accounts for any leading content such as a checkbox.
        /// </summary>
        private static double AvailableWidth(TextBlock tb)
        {
            var list = FindAncestor<ItemsControl>(tb);
            if (list != null && list.ActualWidth > 0)
            {
                try
                {
                    var offset = tb.TransformToAncestor(list).Transform(new Point(0, 0));
                    const double rightInset = 24; // vertical scrollbar + padding breathing room
                    double width = list.ActualWidth - offset.X - rightInset - GetRightReserve(tb);
                    if (width > 0) return width;
                }
                catch
                {
                    // Not parented into the list's visual tree yet; fall through.
                }
            }
            return tb.ActualWidth;
        }

        private static T FindAncestor<T>(DependencyObject from) where T : DependencyObject
        {
            DependencyObject d = VisualTreeHelper.GetParent(from);
            while (d != null && !(d is T)) d = VisualTreeHelper.GetParent(d);
            return d as T;
        }
    }
}
