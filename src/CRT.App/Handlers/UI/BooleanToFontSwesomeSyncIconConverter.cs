using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace CRT;

public sealed class BooleanToFontAwesomeSyncIconConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        bool isEnabled = value as bool? == true;

        return isEnabled
            ? "\uf021" // arrows-rotate / sync
            : "\uf05e"; // ban / disabled
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}