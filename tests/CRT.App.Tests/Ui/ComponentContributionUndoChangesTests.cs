using System.Collections.ObjectModel;
using System.Reflection;
using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Undo changes to [U1]" in the component editor (CRT 4.0 step 1's last slice, 2026-10-25 - the
// plan: "Undo changes to IC1 - one component, back to its published values"). Shown only while the
// component differs from the published board; it puts every row of the component back as ONE step
// of the draft's history (the Drafts tab's Undo takes it back), asks for the board on screen to be
// read again, and shows the component as it now is - or closes, for a component the published board
// does not have.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentContributionUndoChangesTests : IDisposable
{
    private const string BoardIdentity = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    public ComponentContributionUndoChangesTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static BoardData Published() => new()
    {
        Components =
        {
            new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Description = "Video chip" },
            new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Cap" },
        },
    };

    // The draft: U1's description changed, C1 as published, U9 added.
    private static BoardData WriteDraft()
    {
        var draft = new BoardData
        {
            Components =
            {
                new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Description = "Changed in the draft" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Cap" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "New chip" },
            },
        };

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, ComponentContributionUndoChangesTests.BoardIdentity);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, draft);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ComponentContributionUndoChangesTests.BoardIdentity),
            new DraftMarker { BoardKey = ComponentContributionUndoChangesTests.BoardIdentity });

        return draft;
    }

    private ComponentContributionWindow Open(BoardData draft, string label)
    {
        var window = new ComponentContributionWindow
        {
            // Before the load - the load decides the button.
            PublishedBoardOverrideForTests = ComponentContributionUndoChangesTests.Published,
        };

        window.LoadComponent(draft, this.DataRoot, "Commodore 64", "250407", string.Empty, label, ComponentContributionUndoChangesTests.BoardIdentity);
        return window;
    }

    private static Button UndoButton(ComponentContributionWindow window) => window.GetControl<Button>("UndoChangesButton");

    private static ObservableCollection<ContributionComponentRow> ComponentRowsOf(ComponentContributionWindow window) =>
        (ObservableCollection<ContributionComponentRow>)typeof(ComponentContributionWindow)
            .GetField("thisComponentRows", BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(window)!;

    private static ComponentEntry? OnDisk(string label) =>
        DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, ComponentContributionUndoChangesTests.BoardIdentity)!
            .Components.SingleOrDefault(component => component.BoardLabel == label);

    // Shown, red and named for a component the draft changed - and not for one it left alone.
    [Fact]
    public async Task The_button_shows_only_for_a_component_the_draft_changed()
    {
        BoardData draft = WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            ComponentContributionWindow changed = this.Open(draft, "U1");
            await changed.UndoChangesChecked;

            Assert.True(UndoButton(changed).IsVisible);
            Assert.Equal("Undo changes to [U1]", UndoButton(changed).Content);
            Assert.Contains("Danger", UndoButton(changed).Classes);

            ComponentContributionWindow untouched = this.Open(draft, "C1");
            await untouched.UndoChangesChecked;

            Assert.False(UndoButton(untouched).IsVisible);
        });
    }

    // ###########################################################################################
    // Pressed: U1 is back as published on disk - one step, which the draft's Undo takes back - the
    // board is read again, and the window shows U1 as it now is, with nothing left to undo.
    // ###########################################################################################
    [Fact]
    public async Task Undoing_puts_the_component_back_reloads_the_window_and_is_one_step()
    {
        BoardData draft = WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            ComponentContributionWindow window = this.Open(draft, "U1");
            int refreshed = 0;
            window.SetRefreshBoardAfterSave(() => refreshed++);
            await window.UndoChangesChecked;

            await window.UndoChangesAsync();
            await window.UndoChangesChecked;

            Assert.Equal("Video chip", OnDisk("U1")!.Description);
            Assert.NotNull(OnDisk("U9"));
            Assert.Equal(1, refreshed);
            Assert.Equal("Video chip", ComponentRowsOf(window)[0].Description);
            Assert.False(UndoButton(window).IsVisible);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, ComponentContributionUndoChangesTests.BoardIdentity)!;
            Assert.Equal(DraftStepWording.ComponentReverted("U1"), session.UndoDescription);
        });
    }

    // A component the published board does not have goes - and the window, with nothing left to show,
    // closes.
    [Fact]
    public async Task Undoing_a_component_the_draft_added_removes_it_and_closes_the_window()
    {
        BoardData draft = WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            ComponentContributionWindow window = this.Open(draft, "U9");
            await window.UndoChangesChecked;
            Assert.True(UndoButton(window).IsVisible);

            bool closed = false;
            window.Closed += (_, _) => closed = true;
            window.Show();

            await window.UndoChangesAsync();

            Assert.Null(OnDisk("U9"));
            Assert.Equal("Changed in the draft", OnDisk("U1")!.Description);
            Assert.True(closed);
        });
    }
}
