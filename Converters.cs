using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Globalization;
using System.Windows.Media;

namespace PowerpointJabber
{
    class Converters
    {
        public static BoolToSelectedColourConverter boolToSelectedColourConverter = new BoolToSelectedColourConverter();
        public static BoolToVisibilityConverter boolToVisibilityConverter = new BoolToVisibilityConverter();
        public static ReverseBoolToVisibilityConverter reverseBoolToVisibilityConverter = new ReverseBoolToVisibilityConverter();
        public static PenVisibilityConverter penVisibilityConverter = new PenVisibilityConverter();
        public static EraserVisibilityConverter eraserVisibilityConverter = new EraserVisibilityConverter();

        public class BoolToSelectedColourConverter : IValueConverter
        {
            private LinearGradientBrush selectedColourBrush = new LinearGradientBrush
            {
                GradientStops = new GradientStopCollection(
                    new List<GradientStop> 
                    {
                        new GradientStop(new Color{A=255,R=254,G=215,B=169},0.0),
                        new GradientStop(new Color{A=255,R=251,G=181,B=101},0.39),
                        new GradientStop(new Color{A=255,R=250,G=152,B=49},0.4),
                        new GradientStop(new Color{A=255,R=253,G=236,B=166},1.0),
                    }
                ),
                StartPoint = new Point(0, 0),
                EndPoint = new Point(0, 1)
            };
            private SolidColorBrush unselectedColourBrush = Brushes.Transparent;

            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (!(value is bool)) return Brushes.Transparent;
                return (bool)value ? (Brush)selectedColourBrush : (Brush)unselectedColourBrush;
            }
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                throw new NotImplementedException();
            }
        }
        public class PenVisibilityConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (!(value is PowerpointJabber.SimplePenWindow.EditingButton.EditingType)) return false;
                return (PowerpointJabber.SimplePenWindow.EditingButton.EditingType)value == SimplePenWindow.EditingButton.EditingType.Pen ? Visibility.Visible : Visibility.Collapsed;
            }
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                throw new NotImplementedException();
            }
        }
        public class EraserVisibilityConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (!(value is PowerpointJabber.SimplePenWindow.EditingButton.EditingType)) return false;
                return (PowerpointJabber.SimplePenWindow.EditingButton.EditingType)value == SimplePenWindow.EditingButton.EditingType.Eraser ? Visibility.Visible : Visibility.Collapsed;
            }
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                throw new NotImplementedException();
            }
        }
        public class BoolToVisibilityConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                return (bool)value ? Visibility.Visible : Visibility.Collapsed;
            }
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                throw new NotImplementedException();
            }
        }
        public class ReverseBoolToVisibilityConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                if ((bool)value)
                {
                    return Visibility.Collapsed;
                }
                return Visibility.Visible;
            }
            public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                if ((Visibility)value == Visibility.Visible)
                {
                    return false;
                }
                return true;
            }
        }

    }
}
