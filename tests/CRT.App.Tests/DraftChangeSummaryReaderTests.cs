using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftChangeSummaryReader - where "Review changes", Submit and Discard draft read what a draft
// changes (CRT 4.0 step 1's last slice, 2026-10-25): the draft AS EVERY EDITOR SHOWS IT (its
// session's content when one is open - a change the table holds unsaved is in it - else its board
// file), against the published board file with both files' lists, so a picture replaced under its
// own name is said. A board that exists only as a draft is compared with nothing.
//
// In "HeadlessUi", as DraftEditsTests: DraftManager's sessions are shared static state.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftChangeSummaryReaderTests : IDisposable
{
    private const string BoardIdentity = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string PublishedFile => Path.Combine(this.DataRoot, "Commodore", "C64", "250407", DataTreeLayout.BoardFileName);

    private string DraftFile => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftChangeSummaryReaderTests.BoardIdentity);

    private string DraftFolder => DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftChangeSummaryReaderTests.BoardIdentity);

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Board(string description, string imageFile = "") => new()
    {
        Components = { new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Description = description } },
        ComponentImages = imageFile.Length == 0 ? [] : [new ComponentImageEntry { BoardLabel = "U1", Name = "Top", File = imageFile }],
    };

    private DraftChangeSummary Read(bool draftOnly = false) =>
        DraftChangeSummaryReader.Read(this.DataRoot, this.DraftsRoot, DraftChangeSummaryReaderTests.BoardIdentity, draftOnly);

    [Fact]
    public void The_draft_file_is_compared_with_the_published_board_file()
    {
        StoredFiles.WriteBoard(this.PublishedFile, Board("Video chip"));
        StoredFiles.WriteBoard(this.DraftFile, Board("Changed"));

        DraftChangeGroup group = Assert.Single(this.Read().Groups);

        Assert.Equal("U1", group.Label);
        Assert.Equal(["Description changed"], group.Lines);
    }

    // A change held in the draft's session - not saved yet - is in the summary.
    [Fact]
    public void An_open_sessions_unsaved_change_is_what_is_summarised()
    {
        StoredFiles.WriteBoard(this.PublishedFile, Board("Video chip"));
        StoredFiles.WriteBoard(this.DraftFile, Board("Video chip"));

        DraftSession session = DraftManager.Sessions.Get(this.DraftsRoot, DraftChangeSummaryReaderTests.BoardIdentity)!;

        Assert.True(this.Read().IsEmpty);

        session.Apply("Changes to [U1]", content => content with { Board = Board("Typed, not saved") });

        Assert.Equal(["Description changed"], Assert.Single(this.Read().Groups).Lines);
    }

    // A picture replaced under the name it had changes no row - and is still said.
    [Fact]
    public void A_file_replaced_under_its_own_name_is_said()
    {
        StoredFiles.WriteBoard(this.PublishedFile, Board("Video chip", "Component images/u1.png"), StoredFiles.Put(this.DataRoot, "Component images/u1.png", "old picture"));
        StoredFiles.WriteBoard(this.DraftFile, Board("Video chip", "Component images/u1.png"), StoredFiles.Put(this.DraftFolder, "Component images/u1.png", "new picture"));

        Assert.Equal(["File [Component images/u1.png] replaced"], Assert.Single(this.Read().Groups).Lines);
    }

    // A board that exists only as a draft is all new, whatever lies in the data folder.
    [Fact]
    public void A_board_that_exists_only_as_a_draft_is_compared_with_nothing()
    {
        StoredFiles.WriteBoard(this.PublishedFile, Board("Video chip"));
        StoredFiles.WriteBoard(this.DraftFile, Board("Video chip"));

        DraftChangeGroup group = Assert.Single(this.Read(draftOnly: true).Groups);

        Assert.Equal(DraftComponentChange.Added, group.Change);
        Assert.Null(DraftChangeSummaryReader.PublishedBoard(this.DataRoot, DraftChangeSummaryReaderTests.BoardIdentity, draftOnly: true));
        Assert.Equal("Video chip", DraftChangeSummaryReader.PublishedBoard(this.DataRoot, DraftChangeSummaryReaderTests.BoardIdentity, draftOnly: false)!.Components[0].Description);
    }

    [Fact]
    public void Without_a_draft_the_summary_is_empty()
    {
        StoredFiles.WriteBoard(this.PublishedFile, Board("Video chip"));

        Assert.True(this.Read().IsEmpty);
    }
}
