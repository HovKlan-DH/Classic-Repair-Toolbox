using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// "NO TEXT A USER READS CALLS A CONTRIBUTOR "THEIR", "THEM" OR "THEY"" (owner request,
// 2026-10-01) - in the MARKUP too.
//
// The wording classes' tests hold that rule for the sentences built in code, and cannot see a
// string written straight into an .axaml file. Two such strings broke it after the rule was
// written ("... it is the only message they get.", code review, 2026-10-01). So every text
// attribute in the Maintainer tab's and the Drafts tab's markup that mentions a contributor is read
// here, and must not also carry one of the three pronouns.
// ###########################################################################################
public sealed class ContributorPronounMarkupTests
{
    private static readonly Regex TextAttribute = new(
        "\\b(?:Text|PlaceholderText|Content|Header|ToolTip\\.Tip|Title)=\"([^\"]*)\"",
        RegexOptions.Compiled);

    private static readonly Regex Pronoun = new("\\b(their|them|they)\\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    [Fact]
    public void No_markup_text_about_a_contributor_calls_the_contributor_they()
    {
        var offending = new List<string>();

        foreach (string folder in new[] { "Maintainer", "Drafts" })
        {
            string root = ContributorPronounMarkupTests.ResolveTabsFolder(folder);

            foreach (string file in Directory.EnumerateFiles(root, "*.axaml", SearchOption.AllDirectories))
            {
                foreach (Match match in TextAttribute.Matches(File.ReadAllText(file)))
                {
                    string text = match.Groups[1].Value;

                    if (text.Contains("contributor", StringComparison.OrdinalIgnoreCase) && Pronoun.IsMatch(text))
                        offending.Add($"{Path.GetFileName(file)}: \"{text}\"");
                }
            }
        }

        Assert.True(offending.Count == 0, "A contributor called their/them/they in: " + string.Join(" | ", offending));
    }

    // The scan itself: a contributor with a pronoun is caught, either alone is not.
    [Theory]
    [InlineData("Text=\"The contributor sees this - it is the only message they get.\"", true)]
    [InlineData("PlaceholderText=\"Why? This is the only message the contributor receives.\"", false)]
    [InlineData("Text=\"Add them one at a time.\"", false)]
    public void The_scan_finds_a_pronoun_only_beside_a_contributor(string markup, bool expected)
    {
        Match match = TextAttribute.Match(markup);

        Assert.True(match.Success);

        string text = match.Groups[1].Value;
        Assert.Equal(expected, text.Contains("contributor", StringComparison.OrdinalIgnoreCase) && Pronoun.IsMatch(text));
    }

    private static string ResolveTabsFolder(string tab)
    {
        string relative = Path.Combine("src", "CRT.App", "Tabs", tab);
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative)))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Path.Combine(directory.FullName, relative);
    }
}
