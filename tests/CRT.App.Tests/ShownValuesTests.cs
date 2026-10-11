using System.Windows.Input;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ShownValues - whether two list rows show the same thing, every public value compared (code
// review, 2026-10-11: moved out of the Drafts tab's row). The Drafts tab keeps a row it would only
// replace with an identical one, so a click on its button is not thrown away; a value missed here
// would leave a row showing something stale, a value wrongly counted would throw the click away
// again. Tested on rows of its own - TabDraftsTests covers the Drafts tab's list on screen.
// ###########################################################################################
public sealed class ShownValuesTests
{
    private sealed class Command : ICommand
    {
        public event EventHandler? CanExecuteChanged
        {
            add { }
            remove { }
        }

        public bool CanExecute(object? parameter) => true;

        public void Execute(object? parameter)
        {
        }
    }

    // A row as the Drafts tab's are: text, a flag, an accent brush and a command of its own.
    private sealed class Row(string text, bool flag = false, IBrush? accent = null)
    {
        public string Text { get; } = text;
        public bool Flag { get; } = flag;
        public IBrush? Accent { get; } = accent;
        public ICommand Act { get; } = new Command();

        // Not shown - not public.
        internal int Hidden { get; init; }
    }

    private static IBrush Brush(Color color, double opacity = 1) => new ImmutableSolidColorBrush(color, opacity);

    [Fact]
    public void Rows_showing_the_same_values_are_the_same_though_each_has_its_own_command()
    {
        Assert.True(ShownValues.Same(new Row("3 rows changed", true), new Row("3 rows changed", true)));
    }

    [Fact]
    public void A_value_shown_differently_makes_the_rows_differ()
    {
        Assert.False(ShownValues.Same(new Row("3 rows changed"), new Row("4 rows changed")));
        Assert.False(ShownValues.Same(new Row("3 rows changed", true), new Row("3 rows changed", false)));
    }

    // An accent is a new brush each time a row is made - the colour is what is shown.
    [Fact]
    public void Brushes_are_the_same_when_they_paint_the_same_colour()
    {
        Assert.True(ShownValues.Same(new Row("a", accent: Brush(Colors.Orange)), new Row("a", accent: Brush(Colors.Orange))));
        Assert.False(ShownValues.Same(new Row("a", accent: Brush(Colors.Orange)), new Row("a", accent: Brush(Colors.Green))));
        Assert.False(ShownValues.Same(new Row("a", accent: Brush(Colors.Orange)), new Row("a", accent: Brush(Colors.Orange, 0.5))));
        Assert.False(ShownValues.Same(new Row("a", accent: Brush(Colors.Orange)), new Row("a")));
        Assert.True(ShownValues.Same(new Row("a"), new Row("a")));
    }

    [Fact]
    public void A_value_the_row_does_not_show_is_not_compared()
    {
        Assert.True(ShownValues.Same(new Row("a") { Hidden = 1 }, new Row("a") { Hidden = 2 }));
    }
}
