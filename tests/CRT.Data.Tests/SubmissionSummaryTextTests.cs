using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SubmissionSummaryText - what a submission says it changed, now that nobody types it (owner request,
// 2026-11-15: the Submit dialog's "What did you change?" went, "as we do have the "Main contribution"").
// The credit's main contribution in plain words, else the draft's changes as "Review changes" says them -
// and never longer than the server takes.
// ###########################################################################################
public sealed class SubmissionSummaryTextTests
{
    private static DraftChangeSummary ComponentsAdded(int count)
    {
        var published = new BoardData { HardwareName = "Commodore 64", BoardName = "250407" };
        var draft = new BoardData
        {
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Components = [.. Enumerable.Range(1, count).Select(index => new ComponentEntry { BoardLabel = $"U{index}", FriendlyName = "Chip", Category = "IC" })],
        };

        return DraftChangeSummary.Build(new BoardContent(published, [], []), new BoardContent(draft, [], []));
    }

    // The words the maintainers read in the queue and the mail - formatting marks are no words.
    [Fact]
    public void The_credits_main_contribution_is_the_summary_in_plain_words()
    {
        Assert.Equal(
            "Drew the bottom side, from my own board",
            SubmissionSummaryText.For("  Drew the **bottom** side, from *my own* board  ", null));
    }

    // A main contribution of several lines and bullets reads as it was written.
    [Fact]
    public void Its_lines_and_bullets_stay_lines()
    {
        Assert.Equal(
            "Recapped the board:\n" + RichNote.BulletPrefix(0) + "C1\n" + RichNote.BulletPrefix(0) + "C2",
            SubmissionSummaryText.For("Recapped the board:\n- C1\n- C2", null));
    }

    [Fact]
    public void The_main_contribution_is_used_even_when_the_changes_are_known()
    {
        Assert.Equal("Fixed U8's part number", SubmissionSummaryText.For("Fixed U8's part number", ComponentsAdded(3)));
    }

    // ###########################################################################################
    // A draft whose work deserves no credit has no main contribution: its changes say it, so the summary is never
    // empty - an empty one is answered with the server's "summary.missing", asking for a sentence there is no box for.
    // ###########################################################################################
    [Fact]
    public void Without_a_main_contribution_the_changes_say_what_was_changed()
    {
        DraftChangeSummary changes = ComponentsAdded(2);

        Assert.Equal(string.Join("\n", changes.AsLines()), SubmissionSummaryText.For(null, changes));
        Assert.Equal(string.Join("\n", changes.AsLines()), SubmissionSummaryText.For("   ", changes));
    }

    [Fact]
    public void Many_changes_are_listed_up_to_a_few_lines_and_the_rest_counted()
    {
        DraftChangeSummary changes = ComponentsAdded(20);
        Assert.True(changes.AsLines().Count > SubmissionSummaryText.ChangeLinesAtMost);

        string[] lines = SubmissionSummaryText.For(null, changes).Split('\n');

        Assert.Equal(SubmissionSummaryText.ChangeLinesAtMost, lines.Length);
        Assert.StartsWith("... and ", lines[^1], StringComparison.Ordinal);
    }

    [Fact]
    public void With_neither_nothing_is_said()
    {
        Assert.Equal(string.Empty, SubmissionSummaryText.For(null, null));
        Assert.Equal(string.Empty, SubmissionSummaryText.For(string.Empty, DraftChangeSummary.Empty));
    }

    // ###########################################################################################
    // The main contribution has no limit of its own, the server's summary has: a longer one is cut, so the
    // submission is not refused for it (summary.too_long is an error). Held against the server's own rule.
    // ###########################################################################################
    [Fact]
    public void A_long_one_is_cut_to_a_summary_the_server_takes()
    {
        string summary = SubmissionSummaryText.For(new string('x', SubmissionFormat.MaximumSummaryLength + 100), null);

        Assert.Equal(SubmissionFormat.MaximumSummaryLength, summary.Length);
        Assert.EndsWith("...", summary, StringComparison.Ordinal);

        var manifest = new SubmissionManifest
        {
            BoardId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = summary,
        };

        Assert.DoesNotContain(
            SubmissionValidator.Validate(manifest, [], publishedCredits: null),
            finding => finding.Code.StartsWith("summary.", StringComparison.Ordinal));
    }

    // Never half a character: utf8mb4 refuses a lone surrogate, and the submission with it.
    [Fact]
    public void The_cut_never_splits_a_character()
    {
        int beforeEmoji = SubmissionFormat.MaximumSummaryLength - 4;
        string text = new string('a', beforeEmoji) + "\U0001F600" + new string('b', 100);

        string summary = SubmissionSummaryText.For(text, null);

        Assert.Equal(new string('a', beforeEmoji) + "...", summary);
    }
}
