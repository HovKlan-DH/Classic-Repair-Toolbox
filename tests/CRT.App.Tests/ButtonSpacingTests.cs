using System.Globalization;
using System.Xml.Linq;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** BUTTONS SIDE BY SIDE ARE ONE GAP APART, EVERYWHERE (owner request, 2026-10-30: "Please align
// everywhere in UI, where it shows multiple buttons on the same line next to each other, that it has
// the same horizontal spacing, so it looks the same everywhere"). *** It was 8px in most dialogs - "a
// little too much" - and 4, 5, 10 and 12 elsewhere. The gap is ButtonRow.Gap (6).
//
// A row of buttons, in every .axaml under src/: a horizontal StackPanel or a WrapPanel whose children
// are all buttons (a ComboBox among them counts as one) and at least two of them Buttons. Its Spacing or
// ItemSpacing is the gap - or, with neither set and every button spaced by its own margin, the right
// margin is. Two buttons sharing a line in a Grid, columns 0 and 2, have a column of the gap between
// them. Segmented view switches (Segment), the Maintainer tab's screen tabs (ScreenTab), buttons
// drawn as links (LearnMore, *Link) and a formatted note box's row of small icons (FormatTool,
// RichNoteTheme.axaml, 2026-11-07) are no row of buttons.
// ###########################################################################################
public sealed class ButtonSpacingTests
{
    private static readonly HashSet<string> Buttons = ["Button", "ToggleButton", "SplitButton", "DropDownButton", "RepeatButton"];

    private static readonly string Gap = ButtonRow.Gap.ToString(CultureInfo.InvariantCulture);

    [Fact]
    public void The_gap_is_six()
    {
        Assert.Equal(6, ButtonRow.Gap);
    }

    [Fact]
    public void Every_row_of_buttons_in_the_markup_has_the_one_gap()
    {
        var offenders = new List<string>();
        int rows = 0;

        foreach (string file in ButtonSpacingTests.SourceFiles())
        {
            foreach ((bool isRow, string? problem) in ButtonSpacingTests.Check(File.ReadAllText(file)))
            {
                rows += isRow ? 1 : 0;

                if (problem is not null)
                    offenders.Add($"{ButtonSpacingTests.Relative(file)}: {problem}");
            }
        }

        // A parse that found nothing would pass: the app has dozens of rows of buttons.
        Assert.True(rows > 40, $"Only {rows} rows of buttons found - the detector is broken");

        Assert.True(
            offenders.Count == 0,
            $"Buttons side by side are {Gap}px apart everywhere (ButtonRow.Gap; owner request, 2026-10-30):\n" + string.Join("\n", offenders));
    }

    // ---- The detector itself, so a broken one cannot let everything through ---------------------

    [Theory]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="8"><Button /><Button /></StackPanel>""")]
    [InlineData("""<StackPanel Orientation="Horizontal"><Button /><Button /></StackPanel>""")]
    [InlineData("""<WrapPanel ItemSpacing="12"><Button /><Button /><Button /></WrapPanel>""")]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="4"><ComboBox /><Button /><Button /></StackPanel>""")]
    [InlineData("""<WrapPanel><Button Margin="0,0,8,4" /><Button Margin="0,0,8,4" /></WrapPanel>""")]
    [InlineData("""<Grid ColumnDefinitions="*,4,*"><Button Grid.Column="0" /><Button Grid.Column="2" /></Grid>""")]
    public void A_row_with_another_gap_is_found(string markup)
    {
        Assert.Contains(ButtonSpacingTests.Check(ButtonSpacingTests.Wrapped(markup)), found => found.Problem is not null);
    }

    [Theory]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="6"><Button /><Button /></StackPanel>""")]
    [InlineData("""<WrapPanel ItemSpacing="6" LineSpacing="4"><Button /><Button /></WrapPanel>""")]
    [InlineData("""<WrapPanel><Button Margin="0,0,6,4" /><Button Margin="0,0,6,4" /></WrapPanel>""")]
    [InlineData("""<Grid ColumnDefinitions="*,6,*"><Button Grid.Column="0" /><Button Grid.Column="2" /></Grid>""")]
    // Not rows of buttons: one button, a button among other controls, a vertical stack, segments, links.
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="8"><Button /><TextBlock /></StackPanel>""")]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="12"><TextBox /><Button /><Button /></StackPanel>""")]
    [InlineData("""<StackPanel Spacing="8"><Button /><Button /></StackPanel>""")]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="0"><Button Classes="Segment First" /><Button Classes="Segment" /></StackPanel>""")]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="10"><Button Classes="WorklogHeaderLink" /><Button Classes="WorklogHeaderLink" /></StackPanel>""")]
    [InlineData("""<WrapPanel ItemSpacing="4"><Button Classes="ScreenTab" /><Button Classes="ScreenTab" /></WrapPanel>""")]
    [InlineData("""<StackPanel Orientation="Horizontal" Spacing="2"><Button Classes="FormatTool" /><Button Classes="FormatTool" /></StackPanel>""")]
    public void A_row_with_the_gap_or_no_row_at_all_passes(string markup)
    {
        Assert.DoesNotContain(ButtonSpacingTests.Check(ButtonSpacingTests.Wrapped(markup)), found => found.Problem is not null);
    }

    // ---- The rule ---------------------------------------------------------------------------------

    // Each row of buttons in `markup`, with what is wrong with its gap - null when nothing is.
    private static IEnumerable<(bool IsRow, string? Problem)> Check(string markup)
    {
        foreach (XElement element in XDocument.Parse(markup).Descendants())
        {
            string name = element.Name.LocalName;
            List<XElement> children = element.Elements().Where(child => !child.Name.LocalName.Contains('.')).ToList();
            List<XElement> buttons = children.Where(child => ButtonSpacingTests.Buttons.Contains(child.Name.LocalName)).ToList();

            if (buttons.Count < 2 || buttons.Any(ButtonSpacingTests.IsNoRowButton))
                continue;

            if (name == "Grid")
            {
                if (ButtonSpacingTests.GridGap(element, children) is { } problem)
                    yield return (true, problem);

                continue;
            }

            string? attribute = name switch
            {
                "StackPanel" when element.Attribute("Orientation")?.Value == "Horizontal" => "Spacing",
                "WrapPanel" => "ItemSpacing",
                _ => null,
            };

            if (attribute is null || children.Any(child => !ButtonSpacingTests.Buttons.Contains(child.Name.LocalName) && child.Name.LocalName != "ComboBox"))
                continue;

            string? spacing = element.Attribute(attribute)?.Value;

            if (spacing is null && children.All(child => child.Attribute("Margin") is not null))
            {
                IEnumerable<string> rights = children.Select(child => child.Attribute("Margin")!.Value.Split(',').ElementAtOrDefault(2) ?? string.Empty);

                yield return (true, rights.All(right => right.Trim() == ButtonSpacingTests.Gap)
                    ? null
                    : $"<{name}> spaces its buttons by margins of [{string.Join(" ", children.Select(child => child.Attribute("Margin")!.Value))}] - the right margin must be {Gap}");

                continue;
            }

            yield return (true, spacing == ButtonSpacingTests.Gap ? null : $"<{name} {attribute}=\"{spacing ?? "(none)"}\"> holding {buttons.Count} buttons - the gap is {Gap}");
        }
    }

    // Two buttons in a Grid's columns 0 and 2: the column between them is the gap.
    private static string? GridGap(XElement grid, IReadOnlyList<XElement> children)
    {
        string[] columns = (grid.Attribute("ColumnDefinitions")?.Value ?? string.Empty).Split(',');

        if (children.Count != 2 || columns.Length != 3 ||
            children.Select(child => child.Attribute("Grid.Column")?.Value ?? "0").Order().SequenceEqual(["0", "2"]) is false)
        {
            return null;
        }

        return columns[1].Trim() == ButtonSpacingTests.Gap ? null : $"<Grid ColumnDefinitions=\"{grid.Attribute("ColumnDefinitions")!.Value}\"> - the column between its two buttons must be {Gap}";
    }

    private static bool IsNoRowButton(XElement button) =>
        (button.Attribute("Classes")?.Value ?? string.Empty)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(name => name is "Segment" or "ScreenTab" or "LearnMore" or "FormatTool" || name.EndsWith("Link", StringComparison.Ordinal));

    private static string Wrapped(string markup) =>
        $"""<UserControl xmlns="https://github.com/avaloniaui">{markup}</UserControl>""";

    // Every .axaml under src/, build output left out.
    private static IEnumerable<string> SourceFiles() =>
        Directory
            .EnumerateFiles(ButtonSpacingTests.RepositoryPath("src"), "*.axaml", SearchOption.AllDirectories)
            .Where(file =>
            {
                string normalised = file.Replace('\\', '/');
                return !normalised.Contains("/bin/", StringComparison.Ordinal) && !normalised.Contains("/obj/", StringComparison.Ordinal);
            });

    private static string Relative(string file) =>
        Path.GetRelativePath(ButtonSpacingTests.RepositoryPath(string.Empty), file).Replace('\\', '/');

    // Walks up from the test binary to the repository root (the folder holding the solution).
    private static string RepositoryPath(string relative)
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, relative);
    }
}
