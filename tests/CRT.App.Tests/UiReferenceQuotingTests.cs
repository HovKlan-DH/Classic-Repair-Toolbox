using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** A TAB, A MENU OR A BUTTON NAMED IN THE UI IS IN QUOTES (owner request, 2026-11-01: "when there is a
// reference to a tab, a menu or a button then it must be encapsulated with """ - "0 nets used in
// Important KiCad signals" was the example). *** Every string literal in the app and the shared library
// and every text attribute of the markup is held to it, for the names that cannot be anything else:
// each tab, and the parts and buttons listed below. A word that is also ordinary English - "the board's
// KiCad data", "Try again in a moment" - cannot be told apart by a test, and is left to whoever writes it.
//
// A label's own text is no reference: a name that OPENS a literal ("Align KiCad data on [Top]") is
// allowed for the listed names. What is written to the log is not UI.
// ###########################################################################################
public sealed class UiReferenceQuotingTests
{
    // "My Devices" is the tab that was "Workbooks" (2026-11-06); both stay, so a sentence naming either bare fails.
    private const string Tabs = "About|Configuration|Contribute|Feedback|Maintainer|Oscilloscope|Overview|Resources|Schematics|Workbooks|My Devices";

    // A tab named with no quote around it - "the Schematics tab", and the possessive.
    private static readonly Regex UnquotedTab = new(@"(?<![""';\w])(?:" + Tabs + @") tab\b", RegexOptions.Compiled);

    // A part or button named part-way through a text, with no quote before it.
    private static readonly Regex UnquotedName = new(
        @"(?<=\w[ ,:;(]|[,:;(] )(?<![""';])(?:Align KiCad data|Mark components manually|Mark all components automatically from KiCad|KiCad: Files|KiCad: Important signals|Board: Overview|Board: Name|Board: Regions|Board: Categories|Board: Schematic images|Show on board|Run full test suite|Discard draft|Review changes)\b",
        RegexOptions.Compiled);

    private static readonly Regex UnquotedUndo = new(@"(?<![""';\w])Undo (?:takes|brings|in the draft|at the top)\b", RegexOptions.Compiled);

    private static readonly Regex Literal = new(@"""(?:[^""\\\r\n]|\\.)*""", RegexOptions.Compiled);

    private static readonly Regex Attribute = new(@"([\w.:]+)=""([^""]*)""", RegexOptions.Compiled);

    private static readonly Regex NotUi = new(@"Logger\.|CrtLog\.|logger\.|\.Log[A-Z]|nameof\(|RunGuarded", RegexOptions.Compiled);

    private static bool IsOffender(string text) =>
        UnquotedTab.IsMatch(text) || UnquotedName.IsMatch(text) || UnquotedUndo.IsMatch(text);

    [Fact]
    public void Every_tab_menu_and_button_named_in_the_UI_is_in_quotes()
    {
        var offenders = new List<string>();

        foreach (string file in SourceFiles("*.cs"))
        {
            string[] lines = File.ReadAllLines(file);

            for (int index = 0; index < lines.Length; index++)
            {
                string code = WithoutComment(lines[index]);

                if (NotUi.IsMatch(code))
                {
                    continue;
                }

                offenders.AddRange(Literal.Matches(code)
                    .Where(literal => IsOffender(Unescaped(literal.Value)))
                    .Select(literal => $"{Relative(file)}:{index + 1}: {literal.Value}"));

                // A sentence carried over to the next line: a name can stand across the two literals.
                if (index + 1 < lines.Length && CarriedOver(code, WithoutComment(lines[index + 1])) is { } carried && IsOffender(carried))
                {
                    offenders.Add($"{Relative(file)}:{index + 1}: {carried}");
                }
            }
        }

        foreach (string file in SourceFiles("*.axaml"))
        {
            string[] lines = File.ReadAllLines(file);
            bool inComment = false;

            for (int index = 0; index < lines.Length; index++)
            {
                string line = lines[index];
                bool wasInComment = inComment;
                int opens = line.LastIndexOf("<!--", StringComparison.Ordinal);
                int closes = line.LastIndexOf("-->", StringComparison.Ordinal);

                if (opens >= 0 && opens > closes)
                {
                    inComment = true;
                }
                else if (closes >= 0)
                {
                    inComment = false;
                }

                if (wasInComment || opens >= 0)
                {
                    continue;
                }

                offenders.AddRange(Attribute.Matches(line)
                    .Where(attribute => IsText(attribute.Groups[1].Value, attribute.Groups[2].Value) &&
                                        IsOffender(attribute.Groups[2].Value.Replace("&quot;", "\"", StringComparison.Ordinal)))
                    .Select(attribute => $"{Relative(file)}:{index + 1}: {attribute.Value}"));
            }
        }

        Assert.True(offenders.Count == 0, "A tab, a menu or a button named in the UI goes in quotes (owner request, 2026-11-01):\n" + string.Join("\n", offenders));
    }

    // The detector finds what it is for, and leaves what is quoted and a label's own text alone.
    [Theory]
    [InlineData("0 nets used in KiCad: Important signals", true)]
    [InlineData("0 nets used in \"KiCad: Important signals\"", false)]
    [InlineData("Light this net up on the Schematics tab", true)]
    [InlineData("the Schematics tab's picture", true)]
    [InlineData("Light this net up on the \"Schematics\" tab", false)]
    [InlineData("KiCad: Important signals (3)", false)]
    [InlineData("Board: Schematic images", false)]
    [InlineData("under Board: Schematic images on the", true)]
    [InlineData("Align KiCad data on [Top]", false)]
    [InlineData("then press Align KiCad data", true)]
    [InlineData("Undo takes it back", true)]
    [InlineData("\"Undo\" takes it back", false)]
    [InlineData("the board's KiCad data", false)]
    public void The_detector_tells_a_reference_from_a_label(string text, bool found)
    {
        Assert.Equal(found, IsOffender(text));
    }

    // "Sign out on the Maintainer " + "tab" passed for a whole release candidate (code review, 2026-11-01):
    // each literal looked clean alone.
    [Fact]
    public void The_detector_reads_a_sentence_carried_over_to_the_next_line()
    {
        string? carried = CarriedOver("    \"Sign out on the Maintainer \" +", "    \"tab to send with another address.\"");

        Assert.True(IsOffender(carried!));
        Assert.False(IsOffender(CarriedOver("    \"Sign out on the \\\"Maintainer\\\" \" +", "    $\"tab to send.\"")!));
        Assert.Null(CarriedOver("    \"A whole sentence.\";", "    \"Another one.\";"));
        Assert.Null(CarriedOver("    \"Carried over to\" +", "    this.Something + \"else\""));
    }

    private static string Unescaped(string literal) => literal.Replace("\\\"", "\"", StringComparison.Ordinal);

    // The text of the literal ending `code` with a "+" and of the one opening `next`, as one - null when
    // the line carries nothing over.
    private static string? CarriedOver(string code, string next)
    {
        string opening = next.TrimStart().TrimStart('$', '@');

        if (!code.TrimEnd().EndsWith("\" +", StringComparison.Ordinal) || !opening.StartsWith('"'))
        {
            return null;
        }

        MatchCollection before = Literal.Matches(code);
        Match after = Literal.Match(opening);

        return before.Count > 0 && after.Success
            ? Unescaped(before[^1].Value[1..^1]) + Unescaped(after.Value[1..^1])
            : null;
    }

    private static bool IsText(string name, string value) =>
        !value.StartsWith('{') &&
        name is "Text" or "ToolTip.Tip" or "Content" or "Watermark" or "PlaceholderText" or "Title" or "Header" or "AutomationProperties.Name";

    // The line before a "//" that is not inside a string.
    private static string WithoutComment(string line)
    {
        bool inString = false;

        for (int index = 0; index < line.Length - 1; index++)
        {
            char character = line[index];

            if (character == '\\' && inString)
            {
                index++;
                continue;
            }

            if (character == '"')
            {
                inString = !inString;
            }
            else if (!inString && character == '/' && line[index + 1] == '/')
            {
                return line[..index];
            }
        }

        return line;
    }

    // The app and the shared library - the server's words are held by their own tests.
    private static IEnumerable<string> SourceFiles(string pattern) =>
        new[] { "src/CRT.App", "src/CRT.Data" }
            .SelectMany(folder => Directory.EnumerateFiles(RepositoryPath(folder), pattern, SearchOption.AllDirectories))
            .Where(file =>
            {
                string normalised = file.Replace('\\', '/');
                return !normalised.Contains("/bin/", StringComparison.Ordinal) && !normalised.Contains("/obj/", StringComparison.Ordinal);
            });

    private static string Relative(string file) =>
        Path.GetRelativePath(RepositoryPath(string.Empty), file).Replace('\\', '/');

    private static string RepositoryPath(string relative)
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, relative);
    }
}
