using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A changelog version's heading reworded for the "About" tab (owner request, 2026-11-06: "Version <b>3.0.0</b>,
// 2026-October-11" in place of the whole heading bold) - ChangelogVersionHeadings. CHANGELOG.md itself is
// never written; only what is shown changes.
// ###########################################################################################
public sealed class ChangelogVersionHeadingsTests
{
    private static MarkdownHeading Heading(string markdown) =>
        Assert.IsType<MarkdownHeading>(Assert.Single(ChangelogVersionHeadings.Of(GitHubMarkdown.Parse(markdown))));

    private static string TextOf(MarkdownHeading heading) => string.Concat(heading.Spans.Select(span => span.Text));

    [Theory]
    [InlineData("# 3.0.0, 2026-October-11", "3.0.0", ", 2026-October-11")]
    [InlineData("# 4.0.0, 2026-November-?", "4.0.0", ", 2026-November-?")]
    [InlineData("## 4.0.0-beta.2, 2026-December-1", "4.0.0-beta.2", ", 2026-December-1")]
    public void A_version_heading_reads_Version_with_only_the_number_bold(string markdown, string version, string rest)
    {
        MarkdownHeading heading = Heading(markdown);

        Assert.True(heading.BoldAsMarked);
        Assert.Equal(
            [("Version ", MarkdownStyle.None), (version, MarkdownStyle.Bold), (rest, MarkdownStyle.None)],
            heading.Spans.Select(span => (span.Text, span.Style)));
        Assert.Equal($"Version {version}{rest}", TextOf(heading));
    }

    // The level is the heading's own - only its words change.
    [Fact]
    public void The_level_stays()
    {
        Assert.Equal(2, Heading("## 3.1.1, 2026-November-5").Level);
    }

    // Any other heading is drawn as GitHub draws it, bold as a whole: another shape, a version with no date,
    // or one with styling of its own - those words are the owner's, and nothing here guesses at them.
    [Theory]
    [InlineData("# Known issues")]
    [InlineData("# 3.0.0")]
    [InlineData("# Version 3.0.0, 2026-October-11")]
    [InlineData("# **3.0.0**, 2026-October-11")]
    [InlineData("# 3.0, 2026-October-11")]
    public void Any_other_heading_is_left_as_it_is(string markdown)
    {
        MarkdownBlock parsed = Assert.Single(GitHubMarkdown.Parse(markdown));

        Assert.Same(parsed, Assert.Single(ChangelogVersionHeadings.Of([parsed])));
        Assert.False(Assert.IsType<MarkdownHeading>(parsed).BoldAsMarked);
    }

    // Only headings: a line of text shaped as a version is no heading, and stays the very block it was.
    [Fact]
    public void A_paragraph_shaped_as_a_version_stays()
    {
        IReadOnlyList<MarkdownBlock> blocks = GitHubMarkdown.Parse("3.0.0, 2026-October-11");

        Assert.Same(Assert.Single(blocks), Assert.Single(ChangelogVersionHeadings.Of(blocks)));
    }

    // The changelog CRT carries opens with a version, and its heading is of the shape reworded - only the shape
    // is held, its words being the owner's and changing with every release.
    [Fact]
    public void The_changelog_this_build_carries_opens_with_a_reworded_version()
    {
        string changelog = CRT.TabAbout.ReadEmbeddedChangelog()!;
        MarkdownHeading first = ChangelogVersionHeadings.Of(GitHubMarkdown.Parse(changelog)).OfType<MarkdownHeading>().First();

        Assert.True(first.BoldAsMarked);
        Assert.StartsWith("Version ", TextOf(first), StringComparison.Ordinal);
    }
}
