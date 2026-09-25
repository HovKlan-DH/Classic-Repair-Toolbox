using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// WHAT HAPPENS TO THE REST OF THE APPLICATION WHEN THIS WINDOW SAVES.
//
// *** THIS FILE EXISTS BECAUSE OF A REPORTED BUG (2026-09-23). *** A contributor edited one
// component's "Short description", clicked "Save to draft", was told the change was visible on
// the board right away - and then could not see it, either on the board or when reopening the
// same component for editing.
//
// The save itself was never at fault: the workbook on disk was correct. What was missing was
// everything AFTER the write. The board the application renders comes out of DataManager's
// cache, so until that cache is cleared and the board re-read, every surface keeps showing the
// data from before the save.
//
// It was missed because this window used to POST to the contribution server rather than write a
// local draft - nothing on this machine changed, so there was nothing to refresh. Phase 6a
// turned it into a local save and the refresh step did not come with it. The label editor, the
// other local-edit path, has had it all along (TabSchematics.LabelEditor.cs clears the board
// cache and then calls ReloadCurrentBoardFromDisk).
//
// The two tests here are deliberately of different kinds, because the bug had two halves:
//
//   1. The save really does change the workbook on disk - the half that always worked, pinned so
//      a future regression there cannot be misread as this bug coming back.
//   2. A successful save NOTIFIES that the board must be re-read - the half that was missing.
//      Main turns that notification into ReloadCurrentBoardFromDisk, the same entry point the
//      label editor uses.
//
// The notification is asserted through the window's REAL save path rather than by invoking the
// callback directly, which would prove only that a field can be set.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentContributionSaveRefreshTests : IDisposable
{
    private const string ExcelDataFile = "Commodore/C64/250407/Data C64 250407.xlsx";

    // Restores the process-wide drafts root, the same courtesy DraftDriftWindowTests pays in its
    // own Dispose. Without it this class becomes the polluter it was written to survive.
    public void Dispose()
    {
        ComponentContributionSaveRefreshTests.PointDraftManagerAt(string.Empty);
    }

    // ###########################################################################################
    // The reported case, end to end: change ONLY the short description and save.
    //
    // Asserted by reading the draft workbook back off disk rather than by inspecting the window,
    // because "the file now says what I typed" is the thing the contributor went looking for and
    // could not find.
    // ###########################################################################################
    [Fact]
    public async Task Saving_an_edited_short_description_writes_it_to_the_draft_workbook()
    {
        using var workspace = new TempWorkspace();

        await ComponentContributionSaveRefreshTests.SaveEditedDescriptionAsync(
            workspace,
            "Video interface chip - PAL");

        BoardData? draft = DraftWorkbookStore.LoadDraftBoard(
            ComponentContributionSaveRefreshTests.DraftsRootOf(workspace),
            ComponentContributionSaveRefreshTests.ExcelDataFile);

        Assert.NotNull(draft);

        ComponentEntry saved = Assert.Single(
            draft!.Components.FindAll(component =>
                string.Equals(component.BoardLabel, "U1", StringComparison.OrdinalIgnoreCase)));

        Assert.Equal("Video interface chip - PAL", saved.Description);

        // The other component must be untouched - the save is component-scoped, and a window the
        // contributor believes is editing U1 must not rewrite C1.
        Assert.Contains(draft.Components, component =>
            string.Equals(component.BoardLabel, "C1", StringComparison.OrdinalIgnoreCase));
    }

    // ###########################################################################################
    // *** THE REGRESSION TEST FOR THE BUG ITSELF. It fails against the version that shipped, where
    // ApplySubmissionOutcome only re-enabled the Submit button and nothing asked for a reload. ***
    //
    // Main supplies a callback that clears the board cache and reloads; this asserts the window
    // actually calls it on a successful save. Without it the workbook is correct and the screen is
    // stale, which is precisely what was reported.
    // ###########################################################################################
    [Fact]
    public async Task A_successful_save_asks_for_the_board_to_be_re_read()
    {
        using var workspace = new TempWorkspace();

        int refreshCount = 0;

        await ComponentContributionSaveRefreshTests.SaveEditedDescriptionAsync(
            workspace,
            "Changed",
            window => window.SetRefreshBoardAfterSave(() => refreshCount++));

        Assert.Equal(1, refreshCount);
    }

    // ###########################################################################################
    // "Save to draft" then takes the contributor to the Drafts tab (owner request,
    // 2026-09-24): the window reports a landed save to Main, which closes it and switches tab.
    // Driven through the WHOLE click path (SubmitAsync, which the click handler awaits), since that
    // is where the report is made - and exactly once, after the save, not before it.
    // ###########################################################################################
    [Fact]
    public async Task A_successful_Save_to_draft_reports_it_so_the_app_can_go_to_the_Drafts_tab()
    {
        using var workspace = new TempWorkspace();

        int afterSaved = 0;
        int refreshedBeforeIt = -1;
        int refreshCount = 0;

        await ComponentContributionSaveRefreshTests.SaveEditedDescriptionAsync(
            workspace,
            "Changed",
            window =>
            {
                window.SetRefreshBoardAfterSave(() => refreshCount++);
                window.SetAfterSaved(() =>
                {
                    afterSaved++;
                    refreshedBeforeIt = refreshCount;
                });
            },
            throughTheClick: true);

        Assert.Equal(1, afterSaved);

        // The board was already re-read - which is also what shows the Drafts tab when this save
        // created the system's first draft - so the switch has a tab to land on.
        Assert.Equal(1, refreshedBeforeIt);
    }

    // ###########################################################################################
    // *** NOT WHILE THE DRAFTS TAB'S TABLE HOLDS UNSAVED EDITS FOR THIS BOARD (owner
    // request, 2026-09-24). *** The save would make the table's own save refused and its edits
    // lost, so nothing is written, a notice is shown, and the window stays - no refresh, no
    // switch to the Drafts tab.
    // ###########################################################################################
    [Fact]
    public async Task Save_to_draft_is_held_back_while_the_Drafts_table_has_unsaved_edits_for_the_board()
    {
        using var workspace = new TempWorkspace();

        var asked = new List<string>();
        int notices = 0;
        int refreshes = 0;
        int afterSaved = 0;

        await ComponentContributionSaveRefreshTests.SaveEditedDescriptionAsync(
            workspace,
            "Must not be written",
            window =>
            {
                window.SetUnsavedTableEditsCheck(board =>
                {
                    asked.Add(board);
                    return true;
                });
                window.ShowSavingBlockedOverrideForTests = () =>
                {
                    notices++;
                    return Task.CompletedTask;
                };
                window.SetRefreshBoardAfterSave(() => refreshes++);
                window.SetAfterSaved(() => afterSaved++);
            },
            throughTheClick: true);

        // Asked about THIS window's board.
        Assert.Equal([ComponentContributionSaveRefreshTests.ExcelDataFile], asked);
        Assert.Equal(1, notices);
        Assert.Equal(0, refreshes);
        Assert.Equal(0, afterSaved);

        BoardData? draft = DraftWorkbookStore.LoadDraftBoard(
            ComponentContributionSaveRefreshTests.DraftsRootOf(workspace),
            ComponentContributionSaveRefreshTests.ExcelDataFile);
        Assert.Equal(
            "The original description",
            draft!.Components.Single(component => component.BoardLabel == "U1").Description);
    }

    [Fact]
    public async Task Save_to_draft_goes_ahead_when_the_table_has_nothing_unsaved_for_the_board()
    {
        using var workspace = new TempWorkspace();
        int notices = 0;

        await ComponentContributionSaveRefreshTests.SaveEditedDescriptionAsync(
            workspace,
            "Written",
            window =>
            {
                window.SetUnsavedTableEditsCheck(_ => false);
                window.ShowSavingBlockedOverrideForTests = () =>
                {
                    notices++;
                    return Task.CompletedTask;
                };
            },
            throughTheClick: true);

        Assert.Equal(0, notices);

        BoardData? draft = DraftWorkbookStore.LoadDraftBoard(
            ComponentContributionSaveRefreshTests.DraftsRootOf(workspace),
            ComponentContributionSaveRefreshTests.ExcelDataFile);
        Assert.Equal("Written", draft!.Components.Single(component => component.BoardLabel == "U1").Description);
    }

    [Fact]
    public async Task A_save_refused_by_validation_stays_in_the_window()
    {
        // A new component with no board label is refused - the window and its message must stay.
        using var workspace = new TempWorkspace();
        int afterSaved = 0;

        await UiTest.RunAsync(async () =>
        {
            var window = new ComponentContributionWindow();
            window.LoadNewComponent(
                ComponentContributionSaveRefreshTests.PublishedBoard(),
                Path.Combine(workspace.Root, "Data"),
                "C64",
                "250407",
                "PAL",
                ComponentContributionSaveRefreshTests.ExcelDataFile);
            window.SetAfterSaved(() => afterSaved++);

            await ComponentContributionSaveRefreshTests.RunSubmitAsync(window);
        });

        Assert.Equal(0, afterSaved);
    }

    // ###########################################################################################
    // Builds a published board, seeds a real draft from it, opens the window on U1, changes only
    // the description and runs the window's own save.
    //
    // The draft is SEEDED rather than left absent so the test exercises the ordinary case - a
    // draft that already exists, which is what the reporter had by the time they looked.
    // ###########################################################################################
    private static async Task SaveEditedDescriptionAsync(
        TempWorkspace workspace,
        string newDescription,
        Action<ComponentContributionWindow>? configure = null,
        bool throughTheClick = false)
    {
        string dataRoot = Path.Combine(workspace.Root, "Data");
        string draftsRoot = ComponentContributionSaveRefreshTests.DraftsRootOf(workspace);

        Directory.CreateDirectory(dataRoot);
        Directory.CreateDirectory(draftsRoot);

        // ###########################################################################################
        // *** POINTED AT THE WORKSPACE BEFORE ANYTHING READS IT. ***
        //
        // DraftManager is a process-wide static and the save reads DraftManager.DraftsRoot directly
        // (HasDraft, then Edit). Other classes in the "HeadlessUi" collection set and RESET it -
        // DraftDriftWindowTests restores it to the empty string when it disposes - so a root set
        // late, or not at all, leaves the save resolving no draft folder and returning false before
        // it reaches the refresh. That is what made this test pass alone and fail in the full run.
        // ###########################################################################################
        ComponentContributionSaveRefreshTests.PointDraftManagerAt(draftsRoot);

        BoardData published = ComponentContributionSaveRefreshTests.PublishedBoard();

        string publishedPath = Path.Combine(
            dataRoot,
            ComponentContributionSaveRefreshTests.ExcelDataFile.Replace('/', Path.DirectorySeparatorChar));

        Directory.CreateDirectory(Path.GetDirectoryName(publishedPath)!);
        BoardWorkbookWriter.Write(publishedPath, published);

        DraftSeedResult seeded = DraftSeeder.SeedFromPublished(
            draftsRoot,
            dataRoot,
            ComponentContributionSaveRefreshTests.ExcelDataFile,
            published);

        Assert.True(seeded.Created, seeded.Reason);

        bool saved = false;

        // ###########################################################################################
        // *** RunAsync, NOT Run. *** The save awaits BoardDataReader.LoadAsync and Task.Run, and
        // blocking on it inside Run's synchronous body blocks the dispatcher thread itself - the
        // suite hangs outright rather than failing. UiTest.RunAsync keeps pumping, which is exactly
        // what its own header says it exists for.
        // ###########################################################################################
        await UiTest.RunAsync(async () =>
        {
            var window = new ComponentContributionWindow();

            window.LoadComponent(
                published,
                dataRoot,
                "C64",
                "250407",
                "PAL",
                "U1",
                ComponentContributionSaveRefreshTests.ExcelDataFile);

            configure?.Invoke(window);

            // Edit ONLY the description, exactly as reported. The window's own row collection is
            // what its save reads, so changing it here is what typing into the box produces.
            var rows = ComponentContributionSaveRefreshTests.ComponentRowsOf(window);
            Assert.NotEmpty(rows);
            rows[0].Description = newDescription;

            if (throughTheClick)
            {
                await ComponentContributionSaveRefreshTests.RunSubmitAsync(window);
                saved = true;
            }
            else
            {
                saved = await ComponentContributionSaveRefreshTests.RunSaveAsync(window);
            }
        });

        Assert.True(saved, "The window reported that the save failed.");
    }

    private static string DraftsRootOf(TempWorkspace workspace)
    {
        return Path.Combine(workspace.Root, "Drafts");
    }

    private static BoardData PublishedBoard()
    {
        return new BoardData
        {
            RevisionDate = "2026-01-15",
            Components =
            {
                new ComponentEntry
                {
                    BoardLabel = "U1",
                    FriendlyName = "VIC-II",
                    Category = "IC",
                    Region = "PAL",
                    Description = "The original description",
                },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            },
        };
    }

    // The save is private and async - reached the same way the other contribution tests reach this
    // window's internals, because the logic is welded to a Window. It is AWAITED, never blocked on:
    // see the note at the call site.
    // "Save to draft" as the click runs it - validation, save, status, and the after-save report.
    private static async Task RunSubmitAsync(ComponentContributionWindow window)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "SubmitAsync",
            BindingFlags.Instance | BindingFlags.NonPublic);

        await (Task)method!.Invoke(window, null)!;
    }

    private static async Task<bool> RunSaveAsync(ComponentContributionWindow window)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "SaveComponentToDraftAsync",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return await (Task<bool>)method!.Invoke(window, null)!;
    }

    private static System.Collections.ObjectModel.ObservableCollection<ContributionComponentRow>
        ComponentRowsOf(ComponentContributionWindow window)
    {
        var field = typeof(ComponentContributionWindow).GetField(
            "thisComponentRows",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return (System.Collections.ObjectModel.ObservableCollection<ContributionComponentRow>)
            field!.GetValue(window)!;
    }

    // DraftManager is a static singleton, exactly like DataManager and UserSettings, so the test
    // points it at the workspace through its own internal seam rather than calling Load().
    private static void PointDraftManagerAt(string draftsRoot)
    {
        var method = typeof(DraftManager).GetMethod(
            "LoadFrom",
            BindingFlags.Static | BindingFlags.NonPublic);

        method!.Invoke(null, new object?[] { draftsRoot });
    }
}
