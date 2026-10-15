using System.Reflection;
using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The submission's table in the submission panel - TabMaintainer.Table.cs. It took the change
// summary's place behind a button first (owner request, 2026-09-26: "could it instead open in the
// existing right-side panel, just alike it does in the CRT app?"), and is now the panel's only view
// (2026-09-26: "make the table the default first view ... just scrap that information").
//
// *** THE FIRST TEST HERE IS THE ONE THAT WAS MISSING. *** Moving the table into the window made
// the window throw in its constructor - it loads its markup itself, so the field generated for
// the table's x:Name was never filled in - and the application would not have started. Nothing
// built the window, so the suite was green; a render caught it.
//
// The window is BUILT, never shown: its OnOpened restores the real signed-in session and fetches
// the live queue. Private members are reached by reflection, as ExternalTargetLauncherTests does
// in CRT.App.Tests - the logic is welded to a Window.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerTableTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(long id) =>
        new(id, "Commodore/C64/250407", "pending", "Corrected U8.", "c@example.com", null);

    // A submission that changes one component image, against a published board.
    private static ReviewTableData Table() =>
        new(
            Version: 1,
            Published: new SubmissionRows
            {
                ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Name = "Pinout", File = "Commodore/C64/250407/old.png" }]
            },
            Submitted: new SubmissionRows
            {
                ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Name = "Pinout", File = "Commodore/C64/250407/new.png" }]
            });

    private static bool TableShown(TabMaintainer main) => main.FindControl<DockPanel>("TablePanel")!.IsVisible;

    // ###########################################################################################
    // Nothing selected: an empty panel saying so. And NO other view - the summary and the button
    // that swapped the table for it are gone.
    // ###########################################################################################
    [Fact]
    public void The_window_builds_with_nothing_selected_and_no_other_view()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.False(TableShown(main));
            Assert.False(main.IsTableOpen);
            Assert.True(main.FindControl<TextBlock>("NoSubmissionText")!.IsVisible);

            Assert.Null(main.FindControl<Button>("ViewTableButton"));
            Assert.Null(main.FindControl<ScrollViewer>("SummaryScrollViewer"));
        });
    }

    // ###########################################################################################
    // A submission's table and a board's are one table in two places, and share one remembered
    // pick - so a pick in either is handed to the other (code review, 2026-10-04: each wrote the one
    // setting without telling the other, so the next launch opened both on whichever was picked
    // last, and the other table went on showing its own).
    // ###########################################################################################
    [Fact]
    public void A_pill_picked_in_either_Maintainer_table_is_picked_in_the_other_and_remembered()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            var remembered = new List<BoardTableRowKinds>();
            main.UseRememberedChoices(BoardTableRowKinds.None, remembered.Add);

            BoardTableEditor submission = main.TableEditorsForSharedChoices[0];
            BoardTableEditor board = main.TableEditorsForSharedChoices[1];

            submission.Filter = BoardTableRowKinds.Errors;

            Assert.Equal(BoardTableRowKinds.Errors, board.FilterWanted);
            Assert.Equal([BoardTableRowKinds.Errors], remembered);

            board.Filter = BoardTableRowKinds.Added | BoardTableRowKinds.Errors;

            Assert.Equal(BoardTableRowKinds.Added | BoardTableRowKinds.Errors, submission.FilterWanted);
            Assert.Equal(BoardTableRowKinds.Added | BoardTableRowKinds.Errors, remembered[^1]);

            // Handed over, not echoed back: each pick is remembered once.
            Assert.Equal(2, remembered.Count);
        });
    }

    // ###########################################################################################
    // Case 21 of the table's search box (owner request, 2026-10-02): the Maintainer tab's table is
    // the same control as the Drafts tab's, so it has the same search box - and it narrows a
    // submission's rows the same way.
    // ###########################################################################################
    [Fact]
    public void The_Maintainer_tabs_table_has_the_same_search_box()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            main.OpenTableForTests(Row(42), Table());
            BoardTableEditor editor = main.TableEditorForTests;
            TextBox box = editor.GetControl<TextBox>("SearchBox");

            int Shown() => ((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<object>().Count();

            Assert.True(box.IsVisible);
            Assert.Equal(1, Shown());

            box.Text = "zzz";
            editor.ApplyPendingSearchForTests();
            Assert.Equal(0, Shown());

            box.Text = "pinout";
            editor.ApplyPendingSearchForTests();
            Assert.Equal(1, Shown());
        });
    }

    // A submission's table fills the panel.
    [Fact]
    public void A_submission_opens_straight_into_its_table()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.OpenTableForTests(Row(42), Table());

            Assert.True(main.IsTableOpen);
            Assert.True(TableShown(main));
            Assert.True(main.TableEditorForTests.HasTable);
        });
    }

    // ###########################################################################################
    // *** A NEW BOARD'S TOOLTIP DOES NOT SAY "PUBLISHED" (owner report, 2026-09-26: "when there is
    // a NEW system and the maintainer changes some data, then it states 'Published value: (empty)'
    // - yes, it will always be empty"). ***
    //
    // A new board is compared with the SUBMISSION as it arrived (ShowTable's `Published ??
    // Submitted`), so the table HAS a baseline and the default wording named a board that does not
    // exist. The window is what knows this, so this is asserted through the real ShowTable rather
    // than on BoardTableDocument alone - a document test cannot see the window forgetting to say it.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_changed_cell_names_the_submission_rather_than_a_published_value()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            // Nothing of this board is published.
            main.OpenTableForTests(Row(42), new ReviewTableData(
                Version: 1,
                Published: null,
                Submitted: new SubmissionRows
                {
                    ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Name = "Pinout", File = "Manu1/Hardware1/Board1/U8.png" }]
                }));

            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;

            // The baseline IS a board (the submission itself), so HasBaseline cannot be the signal.
            Assert.True(document.HasBaseline);
            Assert.Equal("As submitted", document.BaselineLabel);
        });
    }

    // ###########################################################################################
    // *** A PUBLISHED BOARD'S TOOLTIP NAMES THE BETA SOURCE (owner request, 2026-10-05). *** The
    // server compares a submission with its BETA tree, which can already hold what the stable source
    // does not - and "Published value" showing such a value read as an error.
    // ###########################################################################################
    [Fact]
    public void A_published_boards_changed_cell_names_the_BETA_source_value()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            main.OpenTableForTests(Row(42), Table());

            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
            BoardTableSheet images = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!;
            BoardTableCell changed = Assert.Single(images.Rows).Cells[images.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFile)];

            Assert.Equal("BETA source value", document.BaselineLabel);
            Assert.Equal("BETA source value: Commodore/C64/250407/old.png", changed.ToolTip);
        });
    }

    // Leaving the submission (another one chosen, signing out) closes its table - asking first when
    // it holds unsaved changes, which this one does not.
    [Fact]
    public async Task Leaving_the_submission_closes_its_table()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            main.OpenTableForTests(Row(42), Table());

            Assert.True(await main.CloseTableAsync());

            Assert.False(main.IsTableOpen);
            Assert.False(TableShown(main));
            Assert.False(main.TableEditorForTests.HasTable);
        });
    }

    // ###########################################################################################
    // *** EACH SUBMISSION REOPENS ON THE SHEET LAST LOOKED AT IN IT (owner requests, 2026-09-26:
    // "per board, so if I am in 'Important signals' in one board, then I can navigate to another
    // board, and then it will show the last sheet/tab for that board"). *** One not opened yet
    // starts on its first sheet with a change; a remembered sheet whose tab "Show changes only"
    // hides gives way to the first sheet that has one.
    // ###########################################################################################
    [Fact]
    public async Task Each_submission_reopens_on_the_sheet_last_looked_at_in_it()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            BoardTableEditor editor = main.TableEditorForTests;

            void Choose(string sheet) =>
                editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(sheet)!);

            async Task Visit(long id)
            {
                if (main.IsTableOpen)
                    Assert.True(await main.CloseTableAsync());

                main.OpenTableForTests(Row(id), Table());
            }

            await Visit(42);
            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Choose(BoardWorkbookSchema.SheetKiCadImportantSignals);

            // Not opened before: its first sheet with a change - not 42's sheet.
            await Visit(43);
            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Choose(BoardWorkbookSchema.SheetCredits);

            await Visit(42);
            Assert.Equal(BoardWorkbookSchema.SheetKiCadImportantSignals, editor.CurrentSheet!.Name);

            await Visit(43);
            Assert.Equal(BoardWorkbookSchema.SheetCredits, editor.CurrentSheet!.Name);

            // Credits has nothing to show once the filter is on, so its tab is hidden - 43 then
            // opens on the first sheet that has one. (Picked with no table open, so 43's
            // remembered sheet is still Credits.)
            Assert.True(await main.CloseTableAsync());
            editor.Filter = BoardTableRowFilter.Changes;
            main.OpenTableForTests(Row(43), Table());
            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
        });
    }

    // ###########################################################################################
    // *** A NEW BOARD: ONLY THE MAINTAINER'S OWN CHANGES ARE MARKED (owner decision, 2026-09-26). ***
    // Compared with the submission itself, its rows start white - not all green ("everything is
    // new", so green said nothing), and not with a colour key of a lone "0 Flagged" ("where are the
    // others?"). The change pills picked show no row until the maintainer changes one; an edit is
    // then coloured, and shown.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_table_marks_only_the_maintainers_own_changes()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.OpenTableForTests(Row(42), new ReviewTableData(
                Version: 1,
                Published: null,
                Submitted: new SubmissionRows
                {
                    Schematics =
                    [
                        new BoardSchematicEntry { SchematicName = "1N754A", SchematicImageFile = "Manu1/Hardware1/Board1/1N754A.png" },
                        new BoardSchematicEntry { SchematicName = "1N4001", SchematicImageFile = "Manu1/Hardware1/Board1/1N4001.png" },
                        new BoardSchematicEntry { SchematicName = "1N4148", SchematicImageFile = "Manu1/Hardware1/Board1/1N4148.png" }
                    ]
                }));

            BoardTableEditor editor = main.TableEditorForTests;
            BoardTableSheet sheet = editor.CurrentSheet!;

            Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, sheet.Name);
            Assert.All(sheet.Rows, row => Assert.Equal(BoardTableRowState.Unchanged, row.State));
            Assert.Equal(0, sheet.ChangeCount);

            // The whole colour key, all at nothing - each pill a filter.
            Assert.True(editor.GetControl<Border>("AddedPill").IsVisible);
            Assert.Equal("0", editor.GetControl<TextBlock>("AddedCountText").Text);

            editor.Filter = BoardTableRowFilter.Changes;
            Assert.Empty(((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<BoardTableRow>());

            // The maintainer changes a row: that one is marked, and shown.
            BoardTableRow edited = sheet.Rows[1];
            edited.Cells[BoardWorkbookSchema.BoardSchematics.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColSchematicImageFile)].Text = "Manu1/Hardware1/Board1/1N4001 v2.png";
            editor.RefreshPendingForTests();

            Assert.Equal(BoardTableRowState.Modified, edited.State);
            Assert.Equal(1, sheet.ChangeCount);
        });
    }

    // ###########################################################################################
    // *** A DECISION WAITS FOR THE TABLE TO BE SAVED. *** The modal window made this impossible by
    // being modal; beside the table in the panel, Approve would publish the SAVED version while the
    // maintainer looks at another. Refused, with the reason, and nothing sent.
    // ###########################################################################################
    [Fact]
    public async Task A_decision_waits_while_the_table_has_unsaved_changes()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            main.OpenTableForTests(Row(42), Table());

            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
            BoardTableRow row = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(candidate => !candidate.IsDeleted);
            row.Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";

            Assert.True(main.TableEditorForTests.HasUnsavedChanges);

            await (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(main, [ReviewDecisionKind.Approve])!;

            TextBlock message = main.FindControl<TextBlock>("DecisionMessageText")!;
            Assert.Equal(ReviewTableWording.SaveTableBeforeDeciding, message.Text);
            Assert.True(message.IsVisible);
        });
    }

    // ###########################################################################################
    // *** A REFRESH KEEPS THE SELECTED SUBMISSION. *** Replacing the list used to drop the selection
    // and blank the panel - with the table there, that would have thrown the table away. When the
    // submission LEFT the queue, the panel empties and the table closes: nothing could save it.
    // ###########################################################################################
    [Fact]
    public void A_refresh_keeps_the_selected_submission_and_its_table()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            var queue = (List<ReviewQueueRow>)typeof(TabMaintainer).GetField("thisQueue", Any)!.GetValue(main)!;

            queue.AddRange([Row(41), Row(42), Row(43)]);
            typeof(TabMaintainer).GetField("thisSelectedId", Any)!.SetValue(main, (long?)42);
            main.OpenTableForTests(Row(42), Table());

            object? kept = typeof(TabMaintainer).GetMethod("ApplyQueue", Any)!.Invoke(main, [false]);

            Assert.Equal(42, ((ReviewQueueRow)kept!).Id);
            Assert.Equal(42, main.SelectedQueueRowForTests?.Id);
            Assert.True(main.IsTableOpen);

            // Gone from the queue, found by an explicit refresh (after a decision here) - see
            // TabMaintainerQueueTests for what the queue's own background check does instead.
            queue.RemoveAll(row => row.Id == 42);

            Assert.Null(typeof(TabMaintainer).GetMethod("ApplyQueue", Any)!.Invoke(main, [false]));
            Assert.False(main.IsTableOpen);
            Assert.False(TableShown(main));
            Assert.True(main.FindControl<TextBlock>("NoSubmissionText")!.IsVisible);
        });
    }

    // ###########################################################################################
    // *** SAVING THE TABLE WAITS THE WHOLE WINDOW TOO (2026-09-28). *** The server checks the new
    // content as it would a new submission, which takes a moment on a large board, and "Saving..."
    // in the table's own status line was as easy to miss as the decision's "Working...". Up while
    // the request is in flight, lifted after a refusal, which keeps the table and its changes.
    // ###########################################################################################
    [Fact]
    public async Task The_window_waits_while_the_table_is_saved_and_comes_back_after_a_refusal()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            main.OpenTableForTests(Row(42), Table());

            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
            BoardTableRow row = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(candidate => !candidate.IsDeleted);
            row.Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";

            bool overlayUp = false;
            string? sentence = null;

            BusyOverlay overlay = MaintainerTabHost.AddOverlay(main);

            var server = new AnsweringHttpHandler(_ =>
            {
                overlayUp = overlay.IsVisible && overlay.IsBusy;
                sentence = overlay.Message;
                return AnsweringHttpHandler.Refused();
            });

            typeof(TabMaintainer).GetField("thisClient", Any)!
                .SetValue(main, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

            bool saved = await (Task<bool>)typeof(TabMaintainer).GetMethod("SaveTableAsync", Any)!.Invoke(main, null)!;

            Assert.False(saved);
            Assert.True(overlayUp);
            Assert.Equal(ReviewTableWording.SavingWait, sentence);

            Assert.False(overlay.IsVisible);
            Assert.True(main.IsTableOpen);
            Assert.True(main.TableEditorForTests.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // *** QUITTING CRT ASKS ABOUT THE TABLE'S UNSAVED CHANGES (2026-09-29). *** As a window this
    // cancelled its own Closing; as a tab it answers Main.OnWindowClosing through these two. Cancel
    // keeps CRT open with the table as it was; Discard lets it close. Nothing unsaved, nothing asked.
    // ###########################################################################################
    [Fact]
    public async Task Quitting_CRT_asks_about_unsaved_table_changes_and_cancel_keeps_them()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            var asked = new List<UnsavedTableEditsPrompt>();
            UnsavedTableEditsChoice answer = UnsavedTableEditsChoice.Cancel;
            main.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return answer;
            };

            main.OpenTableForTests(Row(42), Table());

            Assert.False(main.HasUnsavedTableEdits);
            Assert.True(await main.ConfirmLeavingTableAsync());
            Assert.Empty(asked);

            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
            BoardTableRow row = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(candidate => !candidate.IsDeleted);
            row.Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";

            Assert.True(main.HasUnsavedTableEdits);

            Assert.False(await main.ConfirmLeavingTableAsync());
            Assert.Equal([UnsavedTableEditsPrompt.LeavingSubmission], asked);
            Assert.True(main.HasUnsavedTableEdits);

            answer = UnsavedTableEditsChoice.Discard;
            Assert.True(await main.ConfirmLeavingTableAsync());
        });
    }

    // ###########################################################################################
    // *** WHILE CRT HAS TO BE UPDATED, LEAVING OFFERS NO SAVE (code review, 2026-10-09). *** The
    // save goes to the server, which turns this CRT away: offered, it was refused under the tab's
    // cover where nobody could read why, and quitting was silently cancelled every time. So the
    // question is Discard or Cancel - and a Save answered anyway sends nothing and leaves nothing.
    // ###########################################################################################
    [Fact]
    public async Task While_CRT_has_to_be_updated_leaving_the_submission_table_asks_without_a_Save()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            var asked = new List<UnsavedTableEditsPrompt>();
            UnsavedTableEditsChoice answer = UnsavedTableEditsChoice.Save;
            main.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return answer;
            };

            main.OpenTableForTests(Row(42), Table());

            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
            BoardTableRow row = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(candidate => !candidate.IsDeleted);
            row.Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";

            main.ShowUpdateRequired(Handlers.Online.AppUpdateRequiredWording.For(Handlers.Online.AppUpdateArea.Maintainer, "Please update CRT.", null));

            Assert.False(await main.ConfirmLeavingTableAsync());
            Assert.Equal([UnsavedTableEditsPrompt.LeavingUpdateRequired], asked);
            Assert.True(main.HasUnsavedTableEdits);

            answer = UnsavedTableEditsChoice.Discard;
            Assert.True(await main.ConfirmLeavingTableAsync());
        });
    }

    // -----------------------------------------------------------------------------------------
    // Code review, 2026-10-13
    // -----------------------------------------------------------------------------------------

    // The tab in a shown window - a cell's editor needs one - with submission 42's table open and
    // "Pinout" typed over its Name cell, NOT committed: the grid commits only when focus moves
    // inside it, never when a button outside it takes the click.
    private static (Window Window, TabMaintainer Tab) TypedIntoACell()
    {
        var main = new TabMaintainer { RecallSessionOverrideForTests = _ => null };
        var window = new Window { Content = main, Width = 1400, Height = 900 };
        window.Show();
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        // Signed in as far as the screen goes, 42 chosen and its table open - the decision bar's
        // own test's way (TabMaintainerDecisionClickTests).
        ReviewQueueRow row = Row(42);
        typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);
        main.ApplyQueueResponse(new ReviewQueueResponse(true, [row], true));
        typeof(TabMaintainer).GetMethod("ShowSubmission", Any)!.Invoke(main, [row]);
        main.OpenTableForTests(row, Table());
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        BoardTableEditor editor = main.TableEditorForTests;
        DataGrid grid = editor.GetControl<DataGrid>("TableGrid");
        BoardTableSheet images = editor.CurrentSheet!;

        editor.SelectCell(images.Rows.First(row => !row.IsDeleted), BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName));
        grid.Focus();
        Assert.True(grid.BeginEdit());
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        TextBox typing = Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(grid).OfType<TextBox>().Single(box => box.IsFocused);
        typing.Text = "Pinout (typed)";

        // Typed, not committed: the table holds no change yet.
        Assert.True(editor.IsEditingCell);
        Assert.False(main.HasUnsavedTableEdits);

        return (window, main);
    }

    // ###########################################################################################
    // *** APPROVE UNDER A CELL BEING TYPED IN SAYS "SAVE FIRST" (code review, 2026-10-13). *** The
    // grid does not commit when Approve takes the click, so the check saw no change and the
    // approval went out on the version WITHOUT the typed value - which then vanished. The cell is
    // committed first now, and nothing is sent.
    // ###########################################################################################
    [Fact]
    public async Task Approve_with_a_cell_still_in_its_editor_says_save_first_and_sends_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, TabMaintainer main) = TypedIntoACell();

            int requests = 0;
            var server = new AnsweringHttpHandler(_ =>
            {
                requests++;
                return AnsweringHttpHandler.Refused();
            });

            main.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            main.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));
            typeof(TabMaintainer).GetField("thisSelectedId", Any)!.SetValue(main, (long?)42);

            await (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(main, [ReviewDecisionKind.Approve])!;

            Assert.Equal(ReviewTableWording.SaveTableBeforeDeciding, main.FindControl<TextBlock>("DecisionMessageText")!.Text);
            Assert.Equal(0, requests);
            Assert.True(main.HasUnsavedTableEdits);

            typeof(TabMaintainer).GetMethod("StopQueueChecks", Any)!.Invoke(main, null);
            window.Close();
        });
    }

    // ###########################################################################################
    // *** QUITTING CRT UNDER A CELL BEING TYPED IN ASKS (code review, 2026-10-13). *** Main's quit
    // check (CommitCellsAndHasUnsavedTableEdits) and the question itself (ConfirmLeavingTableAsync)
    // both commit the cell first; before, CRT closed without a word and the value was gone.
    // ###########################################################################################
    [Fact]
    public async Task Quitting_with_a_cell_still_in_its_editor_counts_it_and_asks()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, TabMaintainer main) = TypedIntoACell();
            var asked = new List<UnsavedTableEditsPrompt>();
            main.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Cancel;
            };

            Assert.True(main.CommitCellsAndHasUnsavedTableEdits());
            Assert.False(await main.ConfirmLeavingTableAsync());
            Assert.Equal([UnsavedTableEditsPrompt.LeavingSubmission], asked);

            typeof(TabMaintainer).GetMethod("StopQueueChecks", Any)!.Invoke(main, null);
            window.Close();
        });
    }

    // The question itself commits too - installing an update reaches it without Main's check.
    [Fact]
    public async Task Leaving_the_table_with_a_cell_still_in_its_editor_asks()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, TabMaintainer main) = TypedIntoACell();
            var asked = new List<UnsavedTableEditsPrompt>();
            main.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Cancel;
            };

            Assert.False(await main.ConfirmLeavingTableAsync());
            Assert.Equal([UnsavedTableEditsPrompt.LeavingSubmission], asked);
            Assert.True(main.HasUnsavedTableEdits);

            typeof(TabMaintainer).GetMethod("StopQueueChecks", Any)!.Invoke(main, null);
            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE TABLE IS READ-ONLY FOR AN ACCOUNT THAT MAY NOT PUBLISH THE SUBMISSION (code review,
    // 2026-10-13). *** The server takes a change only from an account that may decide it, so an
    // editable table let the maintainer type, press Save and only then be refused. Read-only now,
    // whichever arrives first - the detail or the table - with the reason in the status line.
    // ###########################################################################################
    [Fact]
    public async Task The_table_is_read_only_when_the_account_may_not_publish_the_submission()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(42);

            ReviewSubmissionDetail Detail(bool canPublish) =>
                new(row, canPublish, new ReviewChangeSummaryView(false, []), [], [], null);

            // The table first, then the detail saying no.
            main.OpenTableForTests(row, Table());
            Assert.False(main.TableEditorForTests.IsReadOnly);

            main.ShowDetail(Detail(canPublish: false));
            Assert.True(main.TableEditorForTests.IsReadOnly);
            Assert.Equal(ApprovalWording.NotAMaintainer, main.FindControl<TextBlock>("ApprovalStatusText")!.Text);

            // The detail read again, saying yes - a pool changed meanwhile.
            main.ShowDetail(Detail(canPublish: true));
            Assert.False(main.TableEditorForTests.IsReadOnly);

            // The detail first, then the table: read-only from the start.
            Assert.True(await main.CloseTableAsync());
            main.ShowDetail(Detail(canPublish: false));
            main.OpenTableForTests(row, Table());
            Assert.True(main.TableEditorForTests.IsReadOnly);

            // Another submission's detail says nothing about this table.
            Assert.True(await main.CloseTableAsync());
            main.OpenTableForTests(Row(43), Table());
            Assert.False(main.TableEditorForTests.IsReadOnly);
        });
    }

    // ###########################################################################################
    // *** A MINUTE CHECK WHILE "LEAVE THE TABLE?" IS UP DOES NOT SPLIT THE LIST FROM THE PANEL (code
    // review, 2026-10-13). *** The maintainer had clicked 43 with 42's table holding changes; while
    // the prompt waited, the check put the list's selection back on 42 - and after Discard the panel
    // opened 43 while the list showed 42. The answer here runs a check before it returns, as a
    // timer tick under the real dialog would.
    // ###########################################################################################
    [Fact]
    public async Task A_queue_check_while_the_leave_prompt_is_up_keeps_the_clicked_submission()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            MaintainerTabHost.AddOverlay(main);

            ReviewQueueRow first = Row(42);
            ReviewQueueRow second = Row(43);
            var queue = new ReviewQueueResponse(true, [first, second], true);

            main.ApplyQueueResponse(queue);
            ListBox list = main.FindControl<ListBox>("QueueList")!;

            ListBoxItem ItemOf(long id) => main.QueueItemsForTests.Single(item => (item.Tag as ReviewQueueRow)?.Id == id);

            // 42 chosen, its table open with a change in it.
            list.SelectedItem = ItemOf(42);
            await Task.Delay(50);
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            main.OpenTableForTests(first, Table());
            BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
            document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(row => !row.IsDeleted)
                .Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";
            Assert.True(main.HasUnsavedTableEdits);

            // The prompt: a minute check arrives while it waits, then Discard.
            main.UnsavedTableEditsAnswerForTests = _ =>
            {
                main.ApplyQueueResponse(queue, background: true);
                return UnsavedTableEditsChoice.Discard;
            };

            list.SelectedItem = ItemOf(43);
            await Task.Delay(50);
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Assert.Equal(43, main.SelectedQueueRowForTests?.Id);
            Assert.Equal((long?)43, (long?)typeof(TabMaintainer).GetField("thisSelectedId", Any)!.GetValue(main));
            Assert.False(main.HasUnsavedTableEdits);
        });
    }
}
