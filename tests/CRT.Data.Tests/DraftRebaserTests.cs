using System;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Bringing a real draft folder up to date (code review, 2026-10-13): DraftRebaser around
// DraftRebase's merge - the workbook, the board's own files, the base and the marker.
//
// Shares "BoardData": the published and draft workbooks are read through BoardDataReader.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftRebaserTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";
    private const string ImageA = "Commodore/C64/250407/Sheet1.png";
    private const string ImageB = "Commodore/C64/250407/Component images/U8.png";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string PublishedWorkbook => DraftBoardSource.PublishedPathOf(this.DataRoot, BoardKey);

    private string DraftWorkbook => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Board(string revision, params ComponentEntry[] components) => new()
    {
        RevisionDate = revision,
        HardwareName = "Commodore 64",
        BoardName = "250407",
        Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = ImageA }],
        Components = components.ToList(),
        ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "Clock", File = ImageB }],
    };

    private static ComponentEntry Component(string label, string description) =>
        new() { BoardLabel = label, Description = description };

    private void Publish(BoardData board, string imageABytes, string imageBBytes)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(this.PublishedWorkbook)!);
        BoardWorkbookWriter.Write(this.PublishedWorkbook, board);
        this.WriteData(ImageA, imageABytes);
        this.WriteData(ImageB, imageBBytes);
        BoardDataReader.ClearAllCache();
    }

    private void WriteData(string relative, string text)
    {
        string full = Path.Combine(this.DataRoot, relative.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, text);
    }

    private string DraftFileText(string relative) =>
        File.ReadAllText(DraftFolderLayout.GetReferencedFilePath(this.DraftsRoot, BoardKey, relative));

    private void Seed()
    {
        DraftSeedResult seeded = DraftSeeder.SeedFromPublishedFile(this.DraftsRoot, this.DataRoot, BoardKey);
        Assert.True(seeded.Created, seeded.Reason);
    }

    [Fact]
    public void A_seeded_draft_keeps_a_base_and_the_published_content_hash()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        DraftBase? based = DraftBaseStore.Load(this.DraftsRoot, BoardKey);
        DraftStatus? status = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);

        Assert.NotNull(based);
        Assert.Equal("2026-09-01", based!.Revision);
        Assert.Equal(["U8"], based.Board.Components.Select(c => c.BoardLabel).ToArray());
        Assert.Equal(2, based.Files.Count);
        Assert.True(status!.HasBase);
        Assert.Equal(DraftBaseStore.PublishedContentHash(this.PublishedWorkbook), status.BaseContentHash);
        Assert.Equal(DraftDriftState.InSync, DraftDriftCheck.Resolve(status, this.PublishedWorkbook));

        // The base is bookkeeping: it is not the contributor's work.
        Assert.True(DraftFolderLayout.IsDraftOnlyFile(DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey)));
    }

    [Fact]
    public void Bringing_a_draft_up_to_date_keeps_the_contributors_work_and_takes_the_update()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU"), Component("U9", "VIC")), "sheet v1", "u8 v1");
        this.Seed();

        // The contributor fixes U8 and replaces the U8 scope picture.
        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, BoardKey, board => Board(
            board.RevisionDate, Component("U8", "6510 CPU"), Component("U9", "VIC"))));
        File.WriteAllText(DraftFolderLayout.GetReferencedFilePath(this.DraftsRoot, BoardKey, ImageB), "u8 MINE");

        // Somebody else's work is published: U9 corrected, C1 added, the schematic re-scanned and
        // the U8 picture replaced too.
        this.Publish(
            Board("2026-10-01", Component("U8", "CPU"), Component("U9", "VIC-II 6569"), Component("C1", "Cap")),
            "sheet v2",
            "u8 theirs");

        DraftStatus? before = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(before, this.PublishedWorkbook));

        DraftRebaseReport report = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, report.Outcome);
        Assert.Empty(report.Conflicts);

        BoardData draft = BoardDataReader.ReadWorkbookUncached(this.DraftWorkbook)!;
        BoardData published = BoardDataReader.ReadWorkbookUncached(this.PublishedWorkbook)!;

        // Against the published board the draft says only what the contributor changed.
        Assert.Equal(1, BoardDataDiffer.CountChanges(published, draft));
        Assert.Equal("6510 CPU", draft.Components.Single(c => c.BoardLabel == "U8").Description);
        Assert.Equal("VIC-II 6569", draft.Components.Single(c => c.BoardLabel == "U9").Description);
        Assert.Contains(draft.Components, c => c.BoardLabel == "C1");

        // A file the contributor left alone follows the update; the one they replaced is theirs.
        Assert.Equal("sheet v2", this.DraftFileText(ImageA));
        Assert.Equal("u8 MINE", this.DraftFileText(ImageB));
        Assert.Equal(1, report.FilesUpdated);

        // The marker now says what the draft is based on, and the warning is gone.
        DraftStatus? after = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal("2026-10-01", after!.BaseRevision);
        Assert.Equal(DraftDriftState.InSync, DraftDriftCheck.Resolve(after, this.PublishedWorkbook));
        Assert.Equal("2026-10-01", DraftBaseStore.Load(this.DraftsRoot, BoardKey)!.Revision);
    }

    // ###########################################################################################
    // *** THE NEW BASE NOT KEPT, THE DRAFT NOT NOTED AS UP TO DATE (code review, 2026-10-13). *** The
    // marker used to record the new published content beside the OLD base, and every later run then
    // answered "already up to date" - measuring the contributor's work against the wrong board for
    // good. Now the warning stays, and running again finishes the job without doubling anything.
    // ###########################################################################################
    [Fact]
    public void When_the_new_base_cannot_be_kept_the_draft_is_not_noted_and_a_second_run_finishes()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();
        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, BoardKey, board => Board(board.RevisionDate, Component("U8", "6510 CPU"))));

        this.Publish(Board("2026-10-01", Component("U8", "CPU"), Component("C1", "Cap")), "sheet v2", "u8 v1");

        // The base is written through a temporary file beside it: a folder in its place stops it.
        string basePath = DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey);
        string blocker = DraftFolderLayout.TemporaryPathFor(basePath);
        Directory.CreateDirectory(blocker);

        DraftRebaseReport first = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.NotNoted, first.Outcome);
        Assert.Equal("2026-09-01", DraftBaseStore.Load(this.DraftsRoot, BoardKey)!.Revision);

        DraftStatus? stillStale = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal("2026-09-01", stillStale!.BaseRevision);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(stillStale, this.PublishedWorkbook));

        // The draft has the update already.
        Assert.Contains(BoardDataReader.ReadWorkbookUncached(this.DraftWorkbook)!.Components, c => c.BoardLabel == "C1");

        Directory.Delete(blocker);

        DraftRebaseReport second = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, second.Outcome);
        Assert.Empty(second.Conflicts);

        BoardData draft = BoardDataReader.ReadWorkbookUncached(this.DraftWorkbook)!;
        Assert.Equal(["U8", "C1"], draft.Components.Select(c => c.BoardLabel).ToArray());
        Assert.Equal("6510 CPU", draft.Components.Single(c => c.BoardLabel == "U8").Description);

        DraftStatus? after = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.InSync, DraftDriftCheck.Resolve(after, this.PublishedWorkbook));
        Assert.Equal("2026-10-01", DraftBaseStore.Load(this.DraftsRoot, BoardKey)!.Revision);
    }

    // ###########################################################################################
    // *** A COPY NOTHING CITES ANY MORE GOES (code review, 2026-10-13). *** The update dropped U8's
    // image row: the draft's untouched copy of the picture stayed, the new base no longer listed it,
    // and the Drafts row counted it as the contributor's change - Submit on, with nothing to send.
    // A copy the contributor changed stays theirs.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_copy_the_update_stopped_citing_is_removed_unless_the_contributor_changed_it(bool contributorChangedIt)
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        string copy = DraftFolderLayout.GetReferencedFilePath(this.DraftsRoot, BoardKey, ImageB);

        if (contributorChangedIt)
            File.WriteAllText(copy, "u8 MINE");

        // The update drops the image row, and the server's tidying removes the picture.
        BoardData withoutImage = Board("2026-10-01", Component("U8", "CPU"));
        withoutImage = new BoardData
        {
            RevisionDate = withoutImage.RevisionDate,
            HardwareName = withoutImage.HardwareName,
            BoardName = withoutImage.BoardName,
            Schematics = withoutImage.Schematics,
            Components = withoutImage.Components,
        };

        this.Publish(withoutImage, "sheet v1", "unused");
        File.Delete(Path.Combine(this.DataRoot, ImageB.Replace('/', Path.DirectorySeparatorChar)));

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey).Outcome);

        Assert.Equal(contributorChangedIt, File.Exists(copy));

        DraftStatus? status = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(contributorChangedIt ? 1 : 0, DraftStatusReader.CountOtherChangesCached(status, this.DraftsRoot).Files);
    }

    // ###########################################################################################
    // *** "I HAVE LOOKED AT THIS" QUIETS A DRAFT JUDGED BY CONTENT TOO (code review, 2026-10-13). ***
    // A draft whose base cannot be used is offered only "I have looked at this" - and the content
    // check never read what it recorded, so the warning could never be dismissed.
    // ###########################################################################################
    [Fact]
    public void Seeing_the_drift_of_a_draft_judged_by_content_quiets_it_until_the_published_board_moves_again()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        this.Publish(Board("2026-10-01", Component("U8", "CPU 6510")), "sheet v2", "u8 v2");

        DraftStatus? stale = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(stale, this.PublishedWorkbook));

        Assert.True(DraftWorkbookStore.MarkDriftSeen(
            this.DraftsRoot, BoardKey, "2026-10-01", DraftBaseStore.PublishedContentHash(this.PublishedWorkbook)));

        DraftStatus? seen = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.InSync, DraftDriftCheck.Resolve(seen, this.PublishedWorkbook));

        // The base is not moved by it - the submission still tells the truth.
        Assert.Equal("2026-09-01", seen!.BaseRevision);

        // A later publish - even the same day - warns again.
        this.Publish(Board("2026-10-01", Component("U8", "CPU 6510"), Component("C1", "Cap")), "sheet v2", "u8 v2");
        DraftStatus? movedAgain = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(movedAgain, this.PublishedWorkbook));
    }

    // A base that does not parse is set aside, so the draft reads as having none - and every
    // surface then offers what can actually be done.
    [Fact]
    public void An_unreadable_base_is_set_aside_so_the_draft_reads_as_having_none()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        string basePath = DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey);
        File.WriteAllText(basePath, "{ not json");

        Assert.Null(DraftBaseStore.Load(this.DraftsRoot, BoardKey));
        Assert.False(DraftBaseStore.Exists(this.DraftsRoot, BoardKey));
        Assert.True(File.Exists(basePath + ".unreadable"));
        Assert.False(DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey)!.HasBase);
    }

    [Fact]
    public void A_second_publish_on_the_same_day_is_still_seen_as_a_change()
    {
        // The revision is a date; two publishes on one day carry the same one. The content is not.
        this.Publish(Board("2026-10-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        this.Publish(Board("2026-10-01", Component("U8", "CPU"), Component("C1", "Cap")), "sheet v1", "u8 v1");

        DraftStatus? status = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);

        Assert.Equal(DraftDriftState.Changed, DraftDriftCheck.Resolve(status, this.PublishedWorkbook));
    }

    [Fact]
    public void A_draft_already_up_to_date_is_left_alone()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();
        DateTime written = File.GetLastWriteTimeUtc(this.DraftWorkbook);

        DraftRebaseReport report = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.AlreadyUpToDate, report.Outcome);
        Assert.Equal(written, File.GetLastWriteTimeUtc(this.DraftWorkbook));
    }

    [Fact]
    public void A_draft_without_a_base_cannot_be_brought_up_to_date_and_is_not_touched()
    {
        // A draft started before CRT kept a base.
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();
        File.Delete(DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey));
        string before = DraftWorkbookStore.Fingerprint(this.DraftsRoot, BoardKey);

        this.Publish(Board("2026-10-01", Component("U8", "CPU 6510")), "sheet v2", "u8 v2");

        Assert.Equal(DraftRebaseOutcome.NoBase, DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey).Outcome);
        Assert.Equal(before, DraftWorkbookStore.Fingerprint(this.DraftsRoot, BoardKey));
        Assert.False(DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey)!.HasBase);
    }

    [Fact]
    public void Seeing_the_drift_of_a_draft_without_a_base_quiets_it_until_the_next_publish()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();
        File.Delete(DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey));

        // A draft from before content hashes were kept compares by revision.
        string markerPath = DraftFolderLayout.GetMarkerPath(this.DraftsRoot, BoardKey);
        DraftMarkerStore.Save(markerPath, DraftMarkerStore.Load(markerPath)! with { BaseContentHash = string.Empty });

        this.Publish(Board("2026-10-01", Component("U8", "CPU 6510")), "sheet v2", "u8 v2");
        DraftStatus? stale = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(stale, this.PublishedWorkbook));

        Assert.True(DraftWorkbookStore.MarkDriftSeen(this.DraftsRoot, BoardKey, "2026-10-01"));

        DraftStatus? seen = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.InSync, DraftDriftCheck.Resolve(seen, this.PublishedWorkbook));
        Assert.Equal("2026-09-01", seen!.BaseRevision); // what the server is told stays true

        this.Publish(Board("2026-11-01", Component("U8", "CPU 6510A")), "sheet v3", "u8 v3");
        DraftStatus? again = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(again, this.PublishedWorkbook));
    }
}
