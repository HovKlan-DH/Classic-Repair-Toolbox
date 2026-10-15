using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** A BUTTON'S COLOUR SAYS WHAT IT DOES (owner decision, 2026-10-13: "Throw-away and danger
// equals this IndianRed, while Accept/save/update is a green'ish button"). ***
//
// Until then IndianRed meant two opposite things: "Discard", "Reject" and "Delete" were red, and so
// were "Save changes" and the Contribute tab's three buttons, red there as "the main action". A
// hobbyist could not learn red as a warning while the button that KEEPS the work was red too. Now
// App.axaml has two classes - Accept (green, Button_Ok_*) and Danger (IndianRed, Button_Cancel_*) -
// and every other button is plain.
//
// What this holds, in every .axaml under src/:
//   - a Button never paints itself with those brushes - its colour is the class, so one rule
//     decides it and a reader of the markup sees which it is;
//   - a style that does use them is one of the two classes (or a state of one, such as staying
//     green under the pointer);
//   - no button is both;
//   - a label that starts with a throwing-away word ("Delete", "Remove", "Discard", ...) wears
//     Danger, and one that starts with a going-ahead word ("Save", "Submit", "Add", ...) wears
//     Accept. Buttons labelled in code are held by their own tests (MaintainerPoolViewTests,
//     WorkbooksBoardPreviewTests).
// ###########################################################################################
public sealed class ButtonColourTests
{
    private static readonly Regex ColourBrush = new(@"\{DynamicResource Button_(Ok|Cancel)_(Bg|Fg|Border)\}", RegexOptions.Compiled);

    // The first words of a label that throws something away or turns it down.
    private static readonly string[] DangerWords =
    [
        "Delete", "Remove", "Discard", "Reject", "Request changes", "Push back", "Withdraw", "Reset",
    ];

    // The first words of a label that goes ahead: saves, sends, adds, creates, applies, publishes.
    private static readonly string[] AcceptWords =
    [
        "Save", "Submit", "Send", "Add", "Insert", "Create", "Apply", "Approve", "Publish", "Accept",
        "Install", "Update", "Rebuild", "Set password", "Change password", "Change address",
    ];

    [Fact]
    public void Every_button_in_the_markup_takes_its_colour_from_the_Accept_or_Danger_class()
    {
        var offenders = new List<string>();

        foreach (string file in ButtonColourTests.SourceFiles("*.axaml"))
        {
            foreach (string problem in ButtonColourTests.ProblemsIn(File.ReadAllText(file)))
                offenders.Add($"{ButtonColourTests.Relative(file)}: {problem}");
        }

        Assert.True(
            offenders.Count == 0,
            "A button's colour does not follow CRT's two button colours - Accept (green) for a button that " +
            "goes ahead, Danger (IndianRed) for one that throws something away, set as a CLASS (owner " +
            "decision, 2026-10-13; App.axaml):\n" + string.Join("\n", offenders));
    }

    // A parse that found nothing would pass the test above: the markup has buttons of both kinds.
    [Fact]
    public void The_markup_has_buttons_of_both_colours()
    {
        var classes = new List<string>();

        foreach (string file in ButtonColourTests.SourceFiles("*.axaml"))
        {
            classes.AddRange(XDocument.Parse(File.ReadAllText(file))
                .Descendants()
                .Where(element => element.Name.LocalName == "Button")
                .SelectMany(ButtonColourTests.ClassesOf));
        }

        Assert.Contains("Accept", classes);
        Assert.Contains("Danger", classes);
    }

    // ---- The detector itself, so a broken one cannot let everything through ---------------------

    [Theory]
    [InlineData("""<Button Content="Save changes" Background="{DynamicResource Button_Cancel_Bg}" />""")]
    [InlineData("""<Button Classes="Accept" Content="Save" Foreground="{DynamicResource Button_Ok_Fg}" />""")]
    [InlineData("""<Button Content="Discard" />""")]
    [InlineData("""<Button Classes="Accept" Content="Delete row" />""")]
    [InlineData("""<Button Content="Save changes" />""")]
    [InlineData("""<Button Classes="Danger" Content="Submit" />""")]
    [InlineData("""<Button Classes="Accept Danger" Content="Close" />""")]
    [InlineData("""<Button><TextBlock Text="Remove image" /></Button>""")]
    [InlineData("""<UserControl.Styles><Style Selector="Button.Main"><Setter Property="Background" Value="{DynamicResource Button_Cancel_Bg}" /></Style></UserControl.Styles>""")]
    public void A_button_whose_colour_does_not_say_what_it_does_is_caught(string markup)
    {
        Assert.NotEmpty(ButtonColourTests.ProblemsIn(ButtonColourTests.Wrapped(markup)));
    }

    [Theory]
    [InlineData("""<Button Classes="Accept" Content="Save changes" />""")]
    [InlineData("""<Button Classes="Danger HeaderAction" Content="Delete workbook" />""")]
    [InlineData("""<Button Content="Cancel" />""")]
    [InlineData("""<Button Content="Clear All" />""")]
    [InlineData("""<Button Content="Reload" />""")]
    [InlineData("""<Button Content="Select folder" />""")]
    [InlineData("""<Button Content="{Binding Name}" />""")]
    [InlineData("""<Button Classes="Accept"><TextBlock Text="Add photo" /></Button>""")]
    [InlineData("""<UserControl.Styles><Style Selector="Button.Accept.Floating:pointerover /template/ ContentPresenter#PART_ContentPresenter"><Setter Property="Background" Value="{DynamicResource Button_Ok_Bg}" /></Style></UserControl.Styles>""")]
    [InlineData("""<Border BorderBrush="{DynamicResource Button_Cancel_Border}" />""")]
    public void A_button_whose_colour_says_what_it_does_is_not_caught(string markup)
    {
        Assert.Empty(ButtonColourTests.ProblemsIn(ButtonColourTests.Wrapped(markup)));
    }

    // ###########################################################################################
    // Every way `markup` breaks the rule, one line each.
    // ###########################################################################################
    private static IReadOnlyList<string> ProblemsIn(string markup)
    {
        XDocument document = XDocument.Parse(markup);
        var problems = new List<string>();

        foreach (XElement button in document.Descendants().Where(element => element.Name.LocalName == "Button"))
        {
            string label = ButtonColourTests.LabelOf(button);
            IReadOnlyList<string> classes = ButtonColourTests.ClassesOf(button);
            bool accept = classes.Contains("Accept");
            bool danger = classes.Contains("Danger");

            foreach (string property in new[] { "Background", "Foreground", "BorderBrush" })
            {
                if (button.Attribute(property)?.Value is string value && ButtonColourTests.ColourBrush.IsMatch(value))
                    problems.Add($"\"{label}\" sets its {property} to {value} - give it the class instead");
            }

            if (accept && danger)
                problems.Add($"\"{label}\" is both Accept and Danger");

            if (ButtonColourTests.StartsWithAny(label, ButtonColourTests.DangerWords) && !danger)
                problems.Add($"\"{label}\" throws something away, so it wears Danger");

            if (ButtonColourTests.StartsWithAny(label, ButtonColourTests.AcceptWords) && !accept)
                problems.Add($"\"{label}\" goes ahead, so it wears Accept");
        }

        foreach (XElement style in document.Descendants().Where(element => element.Name.LocalName == "Style"))
        {
            string selector = style.Attribute("Selector")?.Value ?? string.Empty;

            bool usesColour = style.Descendants()
                .Where(element => element.Name.LocalName == "Setter")
                .Any(setter => setter.Attribute("Value")?.Value is string value && ButtonColourTests.ColourBrush.IsMatch(value));

            if (usesColour && selector.StartsWith("Button", StringComparison.Ordinal)
                && !selector.Contains(".Accept", StringComparison.Ordinal) && !selector.Contains(".Danger", StringComparison.Ordinal))
            {
                problems.Add($"the style \"{selector}\" paints buttons with the Accept or Danger brushes - only those two classes may");
            }
        }

        return problems;
    }

    // The label a person reads: the Content, or the first TextBlock inside. A bound or static
    // label ("{Binding ...}") is not known here.
    private static string LabelOf(XElement button)
    {
        string? label = button.Attribute("Content")?.Value
            ?? button.Descendants().FirstOrDefault(child => child.Name.LocalName == "TextBlock")?.Attribute("Text")?.Value;

        return label is null || label.StartsWith('{') ? string.Empty : label.Trim();
    }

    private static IReadOnlyList<string> ClassesOf(XElement button) =>
        (button.Attribute("Classes")?.Value ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);

    // Whole words: "Set password" is a going-ahead label, "Select folder" is not.
    private static bool StartsWithAny(string label, IEnumerable<string> words) =>
        words.Any(word => label.Equals(word, StringComparison.Ordinal)
            || label.StartsWith(word + " ", StringComparison.Ordinal));

    private static string Wrapped(string markup) =>
        $"""<UserControl xmlns="https://github.com/avaloniaui">{markup}</UserControl>""";

    // Every file of the kind under src/, build output left out.
    private static IEnumerable<string> SourceFiles(string pattern) =>
        Directory
            .EnumerateFiles(ButtonColourTests.RepositoryPath("src"), pattern, SearchOption.AllDirectories)
            .Where(file =>
            {
                string normalised = file.Replace('\\', '/');
                return !normalised.Contains("/bin/", StringComparison.Ordinal) && !normalised.Contains("/obj/", StringComparison.Ordinal);
            });

    private static string Relative(string file) =>
        Path.GetRelativePath(ButtonColourTests.RepositoryPath(string.Empty), file).Replace('\\', '/');

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
