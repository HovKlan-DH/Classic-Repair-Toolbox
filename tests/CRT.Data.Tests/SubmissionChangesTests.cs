using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SubmissionChangeFacts - what a submission changed as it went into BETA (owner request,
// 2026-10-04: "Is it possible to summarize each submission change in a textual form ... to get an
// idea, besides the sometimes vague description from the contributor"). Built at the publish from
// the review's own comparison (ReviewSummary), stored, and worded by the Systems screen's History
// view - so these pin what is kept: the sections that changed, the rows and fields named the way the
// table names them, the files, and the bounds.
// ###########################################################################################
public sealed class SubmissionChangesTests
{
    private static BoardData Board(params ComponentEntry[] components) =>
        new() { Components = [.. components] };

    private static ComponentEntry Component(string label, string partNumber = "906114-01", string region = "") =>
        new() { BoardLabel = label, PartNumber = partNumber, Category = "IC", Region = region };

    [Fact]
    public void Only_the_sections_that_changed_are_kept_with_their_rows_and_fields()
    {
        BoardData before = SubmissionChangesTests.Board(Component("U7"), Component("U8"), Component("R3"));
        // U9 differs from R3 in more than its label - otherwise the two are one row renamed
        // (BoardDataDiffer.PairRenamedRows), which a later test covers.
        BoardData after = SubmissionChangesTests.Board(Component("U7"), Component("U8", partNumber: "251715-01"), Component("U9", partNumber: "6526"));

        SubmissionChanges changes = SubmissionChangeFacts.Build(ReviewSummary.Compare(before, after));

        Assert.False(changes.IsNewSystem);

        SectionChanges components = Assert.Single(changes.Sections);
        Assert.Equal(BoardWorkbookSchema.SheetComponents, components.Section);
        Assert.Equal((1, 1, 1, 0), (components.AddedCount, components.ChangedCount, components.RemovedCount, components.RenamedCount));
        Assert.Equal(["U9"], components.Added);
        Assert.Equal(["R3"], components.Removed);

        ChangedRowFact changed = Assert.Single(components.Changed);
        Assert.Equal("U8", changed.Row);
        Assert.Equal([BoardWorkbookSchema.ColPartNumber], changed.Fields);
    }

    // A row that only changed what identifies it is one row RENAMED - the table's own rule.
    [Fact]
    public void A_row_that_only_changed_its_label_is_one_rename()
    {
        SubmissionChanges changes = SubmissionChangeFacts.Build(ReviewSummary.Compare(
            SubmissionChangesTests.Board(Component("C10")),
            SubmissionChangesTests.Board(Component("C51"))));

        SectionChanges components = Assert.Single(changes.Sections);
        Assert.Equal((0, 0, 0, 1), (components.AddedCount, components.ChangedCount, components.RemovedCount, components.RenamedCount));
        Assert.Equal(new RenamedRowFact("C10", "C51"), Assert.Single(components.Renamed));
    }

    // A row's key as the table words it - "U1 / NTSC", never the separator that draws as a box.
    [Fact]
    public void A_row_is_named_the_way_the_table_names_it()
    {
        SubmissionChanges changes = SubmissionChangeFacts.Build(ReviewSummary.Compare(
            SubmissionChangesTests.Board(Component("U1", region: "PAL")),
            SubmissionChangesTests.Board(Component("U1", region: "PAL"), Component("U1", region: "NTSC"))));

        Assert.Equal(["U1 / NTSC"], Assert.Single(changes.Sections).Added);
        Assert.DoesNotContain(BoardDraftNaturalKeys.Separator, Assert.Single(changes.Sections).Added[0]);
    }

    // ###########################################################################################
    // *** BOUNDED. *** A new system adds every row it has; the history keeps the count and the first
    // ListedPerKind names, so one board cannot make every system's history answer huge.
    // ###########################################################################################
    [Fact]
    public void A_new_system_is_counted_whole_and_named_only_in_part()
    {
        ComponentEntry[] many = Enumerable.Range(1, SubmissionChangeFacts.ListedPerKind + 15).Select(i => Component($"C{i}")).ToArray();

        SubmissionChanges changes = SubmissionChangeFacts.Build(ReviewSummary.Compare(null, SubmissionChangesTests.Board(many)));

        Assert.True(changes.IsNewSystem);

        SectionChanges components = Assert.Single(changes.Sections);
        Assert.Equal(SubmissionChangeFacts.ListedPerKind + 15, components.AddedCount);
        Assert.Equal(SubmissionChangeFacts.ListedPerKind, components.Added.Count);
    }

    // ###########################################################################################
    // The files: new at their path, or replacing other bytes there - asked of the tree BEFORE the
    // publish writes them - and, once the publish has removed them, the ones nothing used any more.
    // The board's own workbook and highlight file are not among them: every publish writes those.
    // ###########################################################################################
    [Fact]
    public void A_file_already_at_its_path_is_replaced_and_one_that_is_not_is_added()
    {
        using var workspace = new TempWorkspace();
        string existing = Path.Combine(workspace.Root, "U8.png");
        File.WriteAllText(existing, "old");

        (IReadOnlyList<string> added, IReadOnlyList<string> replaced) = SubmissionChangeFacts.SplitWrites(
        [
            ("Commodore/C64/250407/U8.png", existing),
            ("Commodore/C64/250407/U9.png", Path.Combine(workspace.Root, "U9.png"))
        ]);

        Assert.Equal(["Commodore/C64/250407/U9.png"], added);
        Assert.Equal(["Commodore/C64/250407/U8.png"], replaced);
    }

    [Fact]
    public void The_removed_files_are_added_after_the_publish_and_the_counts_before_it_are_kept()
    {
        string[] added = Enumerable.Range(1, SubmissionChangeFacts.ListedPerKind + 3).Select(i => $"Board/new{i}.png").ToArray();

        SubmissionChanges built = SubmissionChangeFacts.Build(
            ReviewSummary.Compare(SubmissionChangesTests.Board(Component("U8")), SubmissionChangesTests.Board(Component("U8"))),
            added,
            ["Board/U8.png"]);

        SubmissionChanges done = SubmissionChangeFacts.WithRemovedFiles(built, ["Board/old.pdf", "Board/old.pdf", "Board/older.pdf"]);

        Assert.Equal(SubmissionChangeFacts.ListedPerKind + 3, done.Files.AddedCount);
        Assert.Equal(SubmissionChangeFacts.ListedPerKind, done.Files.Added.Count);
        Assert.Equal(1, done.Files.ReplacedCount);
        Assert.Equal(2, done.Files.RemovedCount);
        Assert.Equal(["Board/old.pdf", "Board/older.pdf"], done.Files.Removed);
    }

    // Nothing changed in the rows and no files: nothing to say - a card then says nothing either.
    [Fact]
    public void A_publish_that_changed_nothing_has_nothing_to_say()
    {
        SubmissionChanges none = SubmissionChangeFacts.Build(ReviewSummary.Compare(
            SubmissionChangesTests.Board(Component("U8")), SubmissionChangesTests.Board(Component("U8"))));

        Assert.Empty(none.Sections);
        Assert.False(SubmissionChangeFacts.HasChanges(none));
        Assert.True(SubmissionChangeFacts.HasChanges(none with { Files = SubmissionChangeFacts.Files(["a.png"], null, null) }));
    }

    // The server stores it as JSON and reads it back for every system screen - the round trip keeps it.
    [Fact]
    public void The_facts_survive_a_round_trip_through_json()
    {
        SubmissionChanges changes = SubmissionChangeFacts.Build(
            ReviewSummary.Compare(SubmissionChangesTests.Board(Component("U8")), SubmissionChangesTests.Board(Component("U8", "x"))),
            ["Board/a.png"],
            null,
            ["Board/b.png"]);

        string json = System.Text.Json.JsonSerializer.Serialize(changes);
        SubmissionChanges back = System.Text.Json.JsonSerializer.Deserialize<SubmissionChanges>(json)!;

        Assert.Equal("U8", back.Sections.Single().Changed.Single().Row);
        Assert.Equal([BoardWorkbookSchema.ColPartNumber], back.Sections.Single().Changed.Single().Fields);
        Assert.Equal(["Board/a.png"], back.Files.Added);
        Assert.Equal(["Board/b.png"], back.Files.Removed);
    }
}
