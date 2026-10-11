using Avalonia.Media;
using System.Linq;
using System.Reflection;
using System.Windows.Input;

namespace CRT
{
    // ###########################################################################################
    // WHETHER TWO LIST ROWS SHOW THE SAME THING - every public value a row's markup can bind to,
    // compared. The Drafts tab keeps a row rather than replace it with one showing the same
    // (TabDrafts.KeepUnchangedRows), so a click on its button is not thrown away. Pure comparison
    // logic, kept out of the tab (code review, 2026-10-11) and tested on rows of its own.
    //
    // Found by reflection, so a value added to a row later is compared too. Commands are left out -
    // each row makes its own, acting on what the row was made for, which the caller compares itself
    // (DraftListItem.ShowsTheSameAs). A brush is the same when it paints the same colour: a row's
    // accent is a new brush each time it is made.
    // ###########################################################################################
    internal static class ShownValues
    {
        public static bool Same<T>(T first, T second)
            where T : class =>
            ShownValues.Of<T>.Properties.All(property =>
                ShownValues.SameValue(property.GetValue(first), property.GetValue(second)));

        private static bool SameValue(object? first, object? second) => (first, second) switch
        {
            (ISolidColorBrush a, ISolidColorBrush b) => a.Color == b.Color && a.Opacity.Equals(b.Opacity),
            _ => Equals(first, second),
        };

        // The values a row of type T shows - read once per type.
        private static class Of<T>
        {
            public static readonly PropertyInfo[] Properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => !typeof(ICommand).IsAssignableFrom(property.PropertyType))
                .ToArray();
        }
    }
}
