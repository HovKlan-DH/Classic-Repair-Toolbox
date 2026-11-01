using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** "COLOR", NEVER "COLOUR", IN ANYTHING A PERSON READS (owner request, 2026-10-30: "always use the
// American 'color' spelling" - said before, and a Redo tooltip still read "Colours of the schematic
// images"). *** Every string literal in src/ and every text attribute of the markup (not a binding, a
// name or a class) is held to it. Code names and comments are not text a person reads.
// ###########################################################################################
public sealed class AmericanSpellingTests
{
    private static readonly Regex British = new(@"colour", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // A string literal on one line - "..." with its escapes; @"..." and $"..." alike.
    private static readonly Regex Literal = new(@"""(?:[^""\\\r\n]|\\.)*""", RegexOptions.Compiled);

    // A markup attribute and its value.
    private static readonly Regex Attribute = new(@"([\w.:]+)=""([^""]*)""", RegexOptions.Compiled);

    [Fact]
    public void No_string_a_person_reads_spells_colour_the_British_way()
    {
        var offenders = new List<string>();

        foreach (string file in AmericanSpellingTests.SourceFiles("*.cs"))
        {
            string[] lines = File.ReadAllLines(file);

            for (int index = 0; index < lines.Length; index++)
            {
                string code = AmericanSpellingTests.WithoutComment(lines[index]);

                offenders.AddRange(Literal.Matches(code)
                    .Where(literal => British.IsMatch(literal.Value))
                    .Select(literal => $"{AmericanSpellingTests.Relative(file)}:{index + 1}: {literal.Value}"));
            }
        }

        foreach (string file in AmericanSpellingTests.SourceFiles("*.axaml"))
        {
            string[] lines = File.ReadAllLines(file);

            for (int index = 0; index < lines.Length; index++)
            {
                offenders.AddRange(Attribute.Matches(lines[index])
                    .Where(attribute => AmericanSpellingTests.IsText(attribute.Groups[1].Value, attribute.Groups[2].Value) && British.IsMatch(attribute.Groups[2].Value))
                    .Select(attribute => $"{AmericanSpellingTests.Relative(file)}:{index + 1}: {attribute.Value}"));
            }
        }

        Assert.True(offenders.Count == 0, "\"color\", not \"colour\", in what a person reads (owner request, 2026-10-30):\n" + string.Join("\n", offenders));
    }

    // The detector finds what it is for, and leaves code names, comments and bindings alone.
    [Theory]
    [InlineData("""    public const string PictureColours = "Colours of the schematic images";""", true)]
    [InlineData("""        Text = $"{count} colours",""", true)]
    [InlineData("""    public const string PictureColours = "Colors of the schematic images";""", false)]
    [InlineData("""    // The colours of a picture - "a quote" in a comment.""", false)]
    [InlineData("""    private void CommitPictureColours()""", false)]
    public void The_detector_reads_literals_only(string line, bool found)
    {
        Assert.Equal(found, Literal.Matches(AmericanSpellingTests.WithoutComment(line)).Any(literal => British.IsMatch(literal.Value)));
    }

    [Theory]
    [InlineData("ToolTip.Tip", "Give every image these colours", true)]
    [InlineData("Text", "Thumbnail colour", true)]
    [InlineData("ToolTip.Tip", "{Binding TraceColourTip}", false)]
    [InlineData("Click", "OnUseColoursForAllClick", false)]
    [InlineData("x:Name", "ColourPicker", false)]
    public void Markup_text_is_held_and_bindings_and_names_are_not(string name, string value, bool found)
    {
        Assert.Equal(found, AmericanSpellingTests.IsText(name, value) && British.IsMatch(value));
    }

    // An attribute whose value a person reads: text, a tooltip, a watermark, a button's content.
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

    // Every file of the kind under src/, build output left out.
    private static IEnumerable<string> SourceFiles(string pattern) =>
        Directory
            .EnumerateFiles(AmericanSpellingTests.RepositoryPath("src"), pattern, SearchOption.AllDirectories)
            .Where(file =>
            {
                string normalised = file.Replace('\\', '/');
                return !normalised.Contains("/bin/", StringComparison.Ordinal) && !normalised.Contains("/obj/", StringComparison.Ordinal);
            });

    private static string Relative(string file) =>
        Path.GetRelativePath(AmericanSpellingTests.RepositoryPath(string.Empty), file).Replace('\\', '/');

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
