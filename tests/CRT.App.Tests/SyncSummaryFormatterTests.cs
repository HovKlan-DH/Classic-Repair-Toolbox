using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for SyncSummaryFormatter - the per-file breakdown written to the log underneath the
// one-line sync summary.
//
// The point of this class is that "10 new, 13 updated" in the log and in the main window banner
// never says WHICH files, so what is pinned down here is that every counted file gets named, that
// each group is headed by its own count (so the breakdown can be checked against the summary line
// at a glance), and that a group with nothing in it produces no header at all.
//
// Files that were already up to date are deliberately neither passed in nor listed - on a normal
// launch that set is the entire data folder.
public class SyncSummaryFormatterTests
{
    private static readonly string[] None = [];

    [Fact]
    public void Each_group_is_headed_by_its_own_count_and_names_its_files()
    {
        string[] newFiles = ["Commodore/C64/250469/new-a.png", "Commodore/C64/250469/new-b.png"];
        string[] updatedFiles = ["Commodore/C64/250469/Board.xlsx"];
        string[] failedFiles = ["Amstrad/CPC464/failed.png"];
        string[] invalidFiles = ["../escape.png"];

        var lines = SyncSummaryFormatter.BuildFileBreakdown(newFiles, updatedFiles, failedFiles, invalidFiles);

        string[] expected =
        [
            "    New [2]:",
            "        [Commodore/C64/250469/new-a.png]",
            "        [Commodore/C64/250469/new-b.png]",
            "    Updated [1]:",
            "        [Commodore/C64/250469/Board.xlsx]",
            "    Failed [1]:",
            "        [Amstrad/CPC464/failed.png]",
            "    Rejected [1]:",
            "        [../escape.png]",
        ];

        Assert.Equal(expected, lines);
    }

    // Empty groups are skipped outright - a normal sync has nothing failed and nothing rejected,
    // and two empty headers under every summary would train the reader to skip the whole block.
    [Fact]
    public void An_empty_group_contributes_no_header()
    {
        string[] updated = ["Commodore/C64/250469/Board.xlsx"];

        var lines = SyncSummaryFormatter.BuildFileBreakdown(None, updated, None, None);

        string[] expected =
        [
            "    Updated [1]:",
            "        [Commodore/C64/250469/Board.xlsx]",
        ];

        Assert.Equal(expected, lines);
    }

    // A sync where nothing changed must add nothing at all under the summary line.
    [Fact]
    public void Nothing_changed_produces_no_lines()
    {
        Assert.Empty(SyncSummaryFormatter.BuildFileBreakdown(None, None, None, None));
    }

    // Nulls are tolerated so a caller that has no list for a group need not invent an empty one.
    [Fact]
    public void A_null_group_is_treated_as_an_empty_one()
    {
        Assert.Empty(SyncSummaryFormatter.BuildFileBreakdown(null, null, null, null));
    }

    // The manifest arrives in whatever order the server wrote it, and the download loop preserves
    // that order. Sorting means the same set of changed files always reads the same way in the log,
    // and files from the same board land next to each other.
    [Fact]
    public void Files_are_sorted_so_the_same_sync_always_reads_the_same_way()
    {
        string[] newFiles =
        [
            "Commodore/C64/250469/z.png",
            "Amstrad/CPC464/a.png",
            "Commodore/C64/250469/a.png",
        ];

        var lines = SyncSummaryFormatter.BuildFileBreakdown(newFiles, None, None, None);

        string[] expected =
        [
            "    New [3]:",
            "        [Amstrad/CPC464/a.png]",
            "        [Commodore/C64/250469/a.png]",
            "        [Commodore/C64/250469/z.png]",
        ];

        Assert.Equal(expected, lines);
    }

    // Case-insensitive sorting, so "Zilog" does not sort ahead of "amstrad" the way an ordinal
    // sort would. Data paths are contributed by hand and their casing is not consistent.
    [Fact]
    public void Sorting_ignores_case()
    {
        string[] newFiles = ["Zilog/z.png", "amstrad/a.png"];

        var lines = SyncSummaryFormatter.BuildFileBreakdown(newFiles, None, None, None);

        Assert.Equal("        [amstrad/a.png]", lines[1]);
        Assert.Equal("        [Zilog/z.png]", lines[2]);
    }

    // A blank manifest entry would otherwise log an empty "[]" line, and the header count would
    // then disagree with the number of files actually named beneath it.
    [Fact]
    public void Blank_file_names_are_dropped_and_do_not_count_towards_the_header()
    {
        string[] newFiles = ["Commodore/C64/250469/a.png", "", "   "];

        var lines = SyncSummaryFormatter.BuildFileBreakdown(newFiles, None, None, None);

        string[] expected =
        [
            "    New [1]:",
            "        [Commodore/C64/250469/a.png]",
        ];

        Assert.Equal(expected, lines);
    }

    // A group of nothing but blanks disappears entirely rather than leaving a "[0]" header.
    [Fact]
    public void A_group_of_only_blank_names_contributes_no_header()
    {
        string[] newFiles = ["", " "];

        Assert.Empty(SyncSummaryFormatter.BuildFileBreakdown(newFiles, None, None, None));
    }

    // Surrounding whitespace is trimmed so the brackets sit tight against the path.
    [Fact]
    public void File_names_are_trimmed()
    {
        string[] newFiles = ["  Commodore/C64/250469/a.png  "];

        var lines = SyncSummaryFormatter.BuildFileBreakdown(newFiles, None, None, None);

        Assert.Equal("        [Commodore/C64/250469/a.png]", lines[1]);
    }
}
