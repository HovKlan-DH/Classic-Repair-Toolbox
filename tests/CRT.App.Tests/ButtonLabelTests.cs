using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** NO BUTTON LABEL HAS "..." (owner request, 2026-10-03: 'rename "Account..." button to
// "Account", and please have a rule for this, as this is not the first time I have seen this - a
// button should never have "..."'). ***
//
// The old desktop convention - "..." on a button that opens a window - is not CRT's: a button says
// what it does, in plain words ("Account", "Browse", "Choose image files"). It kept coming back
// because a habit is not a rule anyone is made to read at the moment of writing a button, so this
// makes it a failing test, the way TestPathLiteralTests does for Windows paths.
//
// What counts as a button's label, in every .axaml under src/:
//   - the Content of any element whose name ends in "Button" (Button, ToggleButton, RepeatButton,
//     SplitButton, DropDownButton, RadioButton);
//   - the Text of a TextBlock or Run INSIDE such an element (a label built from children).
// And in every .cs under src/, a Content set to a string literal - how a button is labelled in
// code. Both the three dots and the single ellipsis character are caught.
//
// Text that is not a button is not touched: a "Loading, please wait..." status line, a text box's
// placeholder, a comment.
// ###########################################################################################
public sealed class ButtonLabelTests
{
    private static readonly Regex CodeContent = new("""Content\s*=\s*\$?@?"(?<text>[^"\r\n]*)""", RegexOptions.Compiled);

    [Fact]
    public void No_button_in_the_markup_has_an_ellipsis_in_its_label()
    {
        var offenders = new List<string>();

        foreach (string file in ButtonLabelTests.SourceFiles("*.axaml"))
        {
            foreach (string label in ButtonLabelTests.EllipsisLabelsInMarkup(File.ReadAllText(file)))
                offenders.Add($"{ButtonLabelTests.Relative(file)}: \"{label}\"");
        }

        Assert.True(
            offenders.Count == 0,
            "A button label ends in \"...\" - CRT's buttons say what they do in plain words, with no ellipsis " +
            "(owner rule, 2026-10-03; CLAUDE.md, \"No button label has ...\"):\n" + string.Join("\n", offenders));
    }

    [Fact]
    public void No_button_labelled_in_code_has_an_ellipsis()
    {
        var offenders = new List<string>();

        foreach (string file in ButtonLabelTests.SourceFiles("*.cs"))
        {
            string[] lines = File.ReadAllLines(file);

            for (int index = 0; index < lines.Length; index++)
            {
                if (ButtonLabelTests.IsEllipsisContentInCode(lines[index]))
                    offenders.Add($"{ButtonLabelTests.Relative(file)}:{index + 1}: {lines[index].Trim()}");
            }
        }

        Assert.True(
            offenders.Count == 0,
            "A Content set in code ends in \"...\" - CRT's buttons say what they do in plain words, with no ellipsis " +
            "(owner rule, 2026-10-03):\n" + string.Join("\n", offenders));
    }

    // ---- The detectors themselves, so a broken one cannot let everything through ----------------

    [Theory]
    [InlineData("""<Button Content="Account..." />""")]
    [InlineData("""<Button Content="Browse…" />""")]
    [InlineData("""<ToggleButton Content="More..." />""")]
    [InlineData("""<Button><StackPanel><TextBlock Text="Files..." /></StackPanel></Button>""")]
    [InlineData("""<Button><TextBlock><Run Text="Open..." /></TextBlock></Button>""")]
    public void A_button_label_with_an_ellipsis_is_caught(string markup)
    {
        Assert.NotEmpty(ButtonLabelTests.EllipsisLabelsInMarkup(ButtonLabelTests.Wrapped(markup)));
    }

    [Theory]
    [InlineData("""<Button Content="Account" />""")]
    [InlineData("""<TextBlock Text="Loading, please wait..." />""")]
    [InlineData("""<TextBox PlaceholderText="Click to select a file..." />""")]
    [InlineData("""<Button><TextBlock Text="Files" /></Button>""")]
    public void Text_that_is_not_a_button_label_or_has_no_ellipsis_is_not_caught(string markup)
    {
        Assert.Empty(ButtonLabelTests.EllipsisLabelsInMarkup(ButtonLabelTests.Wrapped(markup)));
    }

    [Theory]
    [InlineData("""            Content = "Account...",""", true)]
    [InlineData("""        button.Content = $"Delete {rows} rows…";""", true)]
    [InlineData("""            Content = "Account",""", false)]
    [InlineData("""        // "Account..." beside "Sign out" opens the window""", false)]
    [InlineData("""        string text = "Checking data - please wait...";""", false)]
    public void A_content_with_an_ellipsis_set_in_code_is_caught(string line, bool caught)
    {
        Assert.Equal(caught, ButtonLabelTests.IsEllipsisContentInCode(line));
    }

    // ###########################################################################################
    // Every label in `markup` that carries an ellipsis: a *Button's Content, or a TextBlock's or
    // Run's Text inside a *Button. Matched by local name, so the XAML namespaces do not matter.
    // ###########################################################################################
    private static IReadOnlyList<string> EllipsisLabelsInMarkup(string markup)
    {
        XDocument document = XDocument.Parse(markup);
        var labels = new List<string>();

        foreach (XElement button in document.Descendants().Where(ButtonLabelTests.IsButton))
        {
            if (button.Attribute("Content")?.Value is string content)
                labels.Add(content);

            labels.AddRange(button
                .Descendants()
                .Where(child => child.Name.LocalName is "TextBlock" or "Run")
                .Select(child => child.Attribute("Text")?.Value)
                .OfType<string>());
        }

        return labels.Where(ButtonLabelTests.HasEllipsis).ToList();
    }

    private static bool IsButton(XElement element) =>
        element.Name.LocalName.EndsWith("Button", StringComparison.Ordinal);

    // A Content set to a string literal, on a line that is not a comment.
    private static bool IsEllipsisContentInCode(string line)
    {
        if (line.TrimStart().StartsWith("//", StringComparison.Ordinal))
            return false;

        return ButtonLabelTests.CodeContent.Matches(line).Any(match => ButtonLabelTests.HasEllipsis(match.Groups["text"].Value));
    }

    private static bool HasEllipsis(string text) =>
        text.Contains("...", StringComparison.Ordinal) || text.Contains('…');

    private static string Wrapped(string markup) =>
        $"""<UserControl xmlns="https://github.com/avaloniaui">{markup}</UserControl>""";

    // Every file of the kind under src/, build output left out.
    private static IEnumerable<string> SourceFiles(string pattern) =>
        Directory
            .EnumerateFiles(ButtonLabelTests.RepositoryPath("src"), pattern, SearchOption.AllDirectories)
            .Where(file =>
            {
                string normalised = file.Replace('\\', '/');
                return !normalised.Contains("/bin/", StringComparison.Ordinal) && !normalised.Contains("/obj/", StringComparison.Ordinal);
            });

    private static string Relative(string file) =>
        Path.GetRelativePath(ButtonLabelTests.RepositoryPath(string.Empty), file).Replace('\\', '/');

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
