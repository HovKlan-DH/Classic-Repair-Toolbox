using System;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Bringing a real draft folder up to date (code review, 2026-10-13): DraftRebaser around
// DraftRebase's merge - the board file and its file list, the base and the marker. In CRT 4.0
// (2026-10-24) a board names its files and the tree's store keeps their bytes; a draft names the
// same files, and only the bytes the contributor brought are in its own store.
//
// Shares "BoardData": the published and draft workbooks are read through BoardDataReader.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftRebaserTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";
    private const string ImageA = "Sheet1.png";
    private const string ImageB = "Component images/U8.png";

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

    // The published board with the two files it names, their bytes in the tree's store. A board
    // whose rows cite only one of them names only that one.
    private void Publish(BoardData board, string imageABytes, string imageBBytes)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(this.PublishedWorkbook)!);
        BoardWorkbookWriter.Write(this.PublishedWorkbook, board);

        var cited = DraftFiles.CitedNames(board).ToHashSet(StringComparer.OrdinalIgnoreCase);
        BoardDocument.StoredFileRow a = StoredFiles.Put(this.DataRoot, ImageA, imageABytes);
        BoardDocument.StoredFileRow b = StoredFiles.Put(this.DataRoot, ImageB, imageBBytes);
        BoardFiles.Change(this.PublishedWorkbook, document => document.Files = new[] { a, b }.Where(file => cited.Contains(file.Name)).ToList());

        BoardDataReader.ClearAllCache();
    }

    private string DraftFolder => DraftFolderLayout.GetBoardFolder(this.DraftsRoot, BoardKey);

    private string? DraftFilePath(string name) => DraftFileResolver.Resolve(this.DataRoot, this.DraftFolder, name, BoardKey);

    private string DraftFileText(string name) => File.ReadAllText(this.DraftFilePath(name)!);

    // The contributor replaces a picture: the new file is picked into the draft folder under its
    // name, and the draft's next write takes it in.
    private void ReplaceInDraft(string name, string text)
    {
        string arrived = Path.Combine(this.DraftFolder, name.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(arrived)!);
        File.WriteAllText(arrived, text);

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, BoardKey, board => board));
        Assert.False(File.Exists(arrived));
    }

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
        this.ReplaceInDraft(ImageB, "u8 MINE");

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

        // The base is replaced by moving a new file over it: held open, it cannot be (Windows -
        // elsewhere file locks are advisory).
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");
        string basePath = DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey);
        DraftRebaseReport first;

        using (new FileStream(basePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            first = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);
        }

        Assert.Equal(DraftRebaseOutcome.NotNoted, first.Outcome);
        Assert.Equal("2026-09-01", DraftBaseStore.Load(this.DraftsRoot, BoardKey)!.Revision);

        DraftStatus? stillStale = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal("2026-09-01", stillStale!.BaseRevision);
        Assert.Equal(DraftDriftState.OfficialIsNewer, DraftDriftCheck.Resolve(stillStale, this.PublishedWorkbook));

        // The draft has the update already.
        Assert.Contains(BoardDataReader.ReadWorkbookUncached(this.DraftWorkbook)!.Components, c => c.BoardLabel == "C1");

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
    // *** A FILE NOTHING CITES ANY MORE GOES (code review, 2026-10-13). *** The update dropped U8's
    // image row: the draft's untouched picture stayed, the new base no longer listed it, and the
    // Drafts row counted it as the contributor's change - Submit on, with nothing to send. A picture
    // the contributor replaced stays theirs - and so does the row citing it (DraftRebaseFiles).
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_file_the_update_stopped_citing_is_removed_unless_the_contributor_changed_it(bool contributorChangedIt)
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        if (contributorChangedIt)
            this.ReplaceInDraft(ImageB, "u8 MINE");

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

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey).Outcome);

        BoardDocument draft = BoardFiles.ReadDocument(this.DraftWorkbook)!;
        Assert.Equal(contributorChangedIt, draft.Files.Any(file => file.Name == ImageB));
        Assert.Equal(contributorChangedIt, draft.ScopeBaselines.Any(row => row.File == ImageB));

        if (contributorChangedIt)
            Assert.Equal("u8 MINE", this.DraftFileText(ImageB));

        DraftStatus? status = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey);
        Assert.Equal(contributorChangedIt ? 1 : 0, DraftStatusReader.CountOtherChangesCached(status, this.DraftsRoot).Files);
    }

    // ###########################################################################################
    // *** "YOUR VERSION IS KEPT" INCLUDES THE FILE THE ROW NAMES (code review, 2026-11-02). *** The
    // contributor edits a picture row's note, and the update removes that row and its file. The row
    // stayed, with a conflict line saying so - while the file went, as one "left as copied, removed by
    // the update": the picture blank, and Submit refusing file.missing. The bytes lie in the tree's
    // store, which the next cleanup empties, so they are copied into the draft's own.
    // ###########################################################################################
    [Fact]
    public void A_row_kept_against_the_update_keeps_its_file_after_the_tree_lets_it_go()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, BoardKey, board => DraftRebaserTests.WithImages(
            board,
            [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "Clock", File = ImageB, Note = "Measured with the lid off" }])));

        this.Publish(DraftRebaserTests.WithImages(Board("2026-10-01", Component("U8", "CPU")), []), "sheet v1", "unused");

        DraftRebaseReport report = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, report.Outcome);
        Assert.Contains("Your version is kept", Assert.Single(report.Conflicts).What, StringComparison.Ordinal);

        this.TreeLetsGoOf(ImageB, "u8 v1");

        Assert.Contains(BoardFiles.ReadDocument(this.DraftWorkbook)!.ScopeBaselines, row => row.File == ImageB);
        Assert.Equal("u8 v1", this.DraftFileText(ImageB));
    }

    // The same for a row the contributor ADDED citing that file: the base's row was left alone and goes
    // with the update, but the new one still names the file.
    [Fact]
    public void A_row_the_contributor_added_citing_a_file_the_update_removed_keeps_that_file()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, BoardKey, board => DraftRebaserTests.WithBoardFile(board, ImageB)));

        this.Publish(DraftRebaserTests.WithImages(Board("2026-10-01", Component("U8", "CPU")), []), "sheet v1", "unused");

        DraftRebaseReport report = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, report.Outcome);

        this.TreeLetsGoOf(ImageB, "u8 v1");

        Assert.Equal("u8 v1", this.DraftFileText(ImageB));
    }

    // Its bytes gone from the tree already (the cleanup at launch got there first): nothing can keep
    // them, and the conflict line says so rather than leaving a blank picture to be found at Submit.
    [Fact]
    public void A_kept_rows_file_whose_bytes_are_gone_already_is_said()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, BoardKey, board => DraftRebaserTests.WithImages(
            board,
            [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "Clock", File = ImageB, Note = "Measured with the lid off" }])));

        this.Publish(DraftRebaserTests.WithImages(Board("2026-10-01", Component("U8", "CPU")), []), "sheet v1", "unused");
        this.TreeLetsGoOf(ImageB, "u8 v1");

        DraftRebaseReport report = DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey);

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, report.Outcome);

        DraftRebaseConflict gone = Assert.Single(report.Conflicts, conflict => conflict.Section == DraftRebaseFiles.FilesSection);
        Assert.Equal(ImageB, gone.Row);
        Assert.Contains("no longer on this computer", gone.What, StringComparison.Ordinal);

        // No entry, as for any name a row cites with no file anywhere: the board's checks say it is missing.
        Assert.DoesNotContain(BoardFiles.ReadDocument(this.DraftWorkbook)!.Files, file => file.Name == ImageB);
        Assert.Null(this.DraftFilePath(ImageB));
    }

    // Nothing is copied for a row the contributor left alone: it goes with the update, as before.
    [Fact]
    public void A_file_nothing_cites_any_more_is_not_copied_into_the_draft()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        this.Publish(DraftRebaserTests.WithImages(Board("2026-10-01", Component("U8", "CPU")), []), "sheet v1", "unused");

        Assert.Empty(DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey).Conflicts);
        Assert.False(Directory.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.FilesFolder)));
    }

    // ###########################################################################################
    // *** A NAME IN THE BASE CAN SAY ANYTHING (code review, 2026-11-02), Security. *** ".crt-base.json"
    // travels with a draft folder somebody hands over, and "is a file of this name lying in the draft?"
    // asked File.Exists of the name joined to the folder: a rooted name replaces the folder, so a UNC
    // name made CRT contact that machine - which hands over the Windows sign-in's hash. Such a name is no
    // file of the draft: its file is not taken for the contributor's replacement.
    // ###########################################################################################
    [Fact]
    public void A_rooted_name_in_the_base_is_never_taken_for_a_file_lying_in_the_draft()
    {
        this.Publish(Board("2026-09-01", Component("U8", "CPU")), "sheet v1", "u8 v1");
        this.Seed();

        string elsewhere = this.thisWorkspace.Path_("Elsewhere", "planted.png");
        Directory.CreateDirectory(Path.GetDirectoryName(elsewhere)!);
        File.WriteAllText(elsewhere, "not the draft's");

        BoardDocument.StoredFileRow rooted = StoredFiles.EntryFor(elsewhere.Replace('\\', '/'), "not the draft's");
        var row = new BoardDocument.BoardFileRow { Category = "Resources", Name = "Planted", File = rooted.Name };

        BoardDocument based = DraftBaseStore.LoadDocument(this.DraftsRoot, BoardKey)!;
        based.BoardFiles.Add(row);
        based.Files.Add(rooted);
        DraftBaseStore.SaveDocument(this.DraftsRoot, BoardKey, based);

        BoardFiles.Change(this.DraftWorkbook, document =>
        {
            document.BoardFiles.Add(row);
            document.Files.Add(rooted);
        });

        this.Publish(Board("2026-10-01", Component("U8", "CPU"), Component("C1", "Cap")), "sheet v1", "u8 v1");

        Assert.Equal(DraftRebaseOutcome.BroughtUpToDate, DraftRebaser.Rebase(this.DraftsRoot, this.DataRoot, BoardKey).Outcome);

        // Left as copied, and removed by the update: it goes. Taken for a file the contributor replaced,
        // its row was held as the contributor's work.
        Assert.DoesNotContain(BoardFiles.ReadDocument(this.DraftWorkbook)!.BoardFiles, file => file.File == rooted.Name);
    }

    // What the server's tidying does once nothing lists a file: its bytes leave the tree's store.
    private void TreeLetsGoOf(string name, string content)
    {
        BoardDocument.StoredFileRow file = StoredFiles.EntryFor(name, content);
        File.Delete(Path.Combine(this.DataRoot, DataTreeLayout.StoredFilePath(file.Sha256, name).Replace('/', Path.DirectorySeparatorChar)));
        BoardDataReader.ClearAllCache();
    }

    private static BoardData WithImages(BoardData board, List<ComponentImageEntry> images) => new()
    {
        RevisionDate = board.RevisionDate,
        HardwareName = board.HardwareName,
        BoardName = board.BoardName,
        Schematics = board.Schematics,
        Components = board.Components,
        ComponentImages = images,
    };

    private static BoardData WithBoardFile(BoardData board, string file) => new()
    {
        RevisionDate = board.RevisionDate,
        HardwareName = board.HardwareName,
        BoardName = board.BoardName,
        Schematics = board.Schematics,
        Components = board.Components,
        ComponentImages = board.ComponentImages,
        BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "The same picture", File = file }],
    };

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
