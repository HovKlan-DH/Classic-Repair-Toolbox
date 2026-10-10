using Avalonia;
using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "CRT has to be updated" on the Drafts and Maintainer tabs, as CRT's window drives it (owner
// request, 2026-10-09: "When an API diff requires for the CRT app to be updated, then for both the
// tabs, "Draft" and "Maintainer" put a fullpage modal (or alike) in there, that cannot be closed").
// Main.UpdateRequired.cs: the API revision at launch covers both tabs; a 426 covers its own tab and
// asks the revision again.
//
// Main is BUILT, never started (MainWindowTests' rule): the health question is answered by
// ServerApiRevisionOverrideForTests and the button by UpdateRequiredActionOverrideForTests, so
// nothing reaches the network, GitHub or a browser. UserSettings and the workbook folder point at
// temp files, as in MainWindowTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainUpdateRequiredTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MainUpdateRequiredTests()
    {
        this.RedirectToTemp();
    }

    public void Dispose()
    {
        this.RedirectToTemp();
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private void RedirectToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

    // A window whose health question answers `serverRevision`, counting how often it is asked.
    private static (CRT.Main Window, Func<int> Asked) BuildWindow(int? serverRevision, string? pendingVersion = null)
    {
        var window = new CRT.Main();
        int asked = 0;

        window.ServerApiRevisionOverrideForTests = () =>
        {
            asked++;
            return Task.FromResult(serverRevision);
        };

        window.PendingVersionOverrideForTests = () => pendingVersion;

        return (window, () => asked);
    }

    // ###########################################################################################
    // *** THE API DIFF COVERS BOTH TABS. *** A server serving a higher revision than this CRT was
    // built for turns it away from submissions and from the Maintainer tab alike - both are covered
    // before anybody opens them, in the server's own sentence.
    // ###########################################################################################
    [Fact]
    public async Task A_server_on_a_higher_api_revision_covers_both_tabs()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + 1);

            await window.AskServerApiRevisionAsync();

            Assert.True(window.TabDrafts.IsUpdateRequiredShown);
            Assert.True(window.TabMaintainer.IsUpdateRequiredShown);

            AppUpdateRequiredView drafts = window.TabDrafts.UpdateRequiredView!;
            Assert.Equal(AppUpdateRequiredWording.Heading, drafts.Heading);
            Assert.Contains("was made for an older version of the server", drafts.Reason, StringComparison.Ordinal);
            Assert.Equal(AppUpdateRequiredWording.WhatItStops(AppUpdateArea.Drafts), drafts.WhatItStops);
            Assert.Equal(AppUpdateRequiredWording.WhatItStops(AppUpdateArea.Maintainer), window.TabMaintainer.UpdateRequiredView!.WhatItStops);
        });
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public async Task A_server_on_this_revision_or_older_covers_nothing(int offset)
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + offset);

            await window.AskServerApiRevisionAsync();

            Assert.False(window.TabDrafts.IsUpdateRequiredShown);
            Assert.False(window.TabMaintainer.IsUpdateRequiredShown);
        });
    }

    // No answer - offline, or a server older than 4.6.0 - is not a refusal.
    [Fact]
    public async Task No_answer_from_the_server_covers_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();
            window.ServerApiRevisionOverrideForTests = () => throw new HttpRequestException("No network.");

            await window.AskServerApiRevisionAsync();

            Assert.False(window.TabDrafts.IsUpdateRequiredShown);
            Assert.False(window.TabMaintainer.IsUpdateRequiredShown);
        });
    }

    // ###########################################################################################
    // *** A REFUSAL COVERS ITS OWN TAB, AND ASKS THE REVISION AGAIN. *** A minimum version set for the
    // Maintainer area turns the Maintainer tab away - in the server's words, which name the version -
    // while the revision still matches, so the contributors' Drafts tab stays open.
    // ###########################################################################################
    [Fact]
    public async Task A_maintainer_refusal_covers_only_the_Maintainer_tab_when_the_revision_still_matches()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, asked) = BuildWindow(ClientVersionContract.ApiRevision);
            string words = ClientVersionContract.Outdated(CrtVersion.Parse("3.0.0"), CrtVersion.Parse("3.2.0")).Message;

            window.ReportApiOutdated(AppUpdateArea.Maintainer, words);
            await SettleAsync(() => asked() == 1);

            Assert.True(window.TabMaintainer.IsUpdateRequiredShown);
            Assert.Equal(words, window.TabMaintainer.UpdateRequiredView!.Reason);
            Assert.False(window.TabDrafts.IsUpdateRequiredShown);
        });
    }

    // A server deployed while CRT runs: the first refusal - here from a submission check - covers the
    // Drafts tab, and the revision asked again covers the Maintainer tab at once, not the first time
    // that tab asks the server. The Drafts tab keeps the server's own words.
    [Fact]
    public async Task A_refusal_from_a_server_on_a_higher_revision_covers_both_tabs()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, asked) = BuildWindow(ClientVersionContract.ApiRevision + 1);

            window.ReportApiOutdated(AppUpdateArea.Drafts, "The server's own sentence.");
            await SettleAsync(() => window.TabMaintainer.IsUpdateRequiredShown);

            Assert.True(window.TabDrafts.IsUpdateRequiredShown);
            Assert.Equal("The server's own sentence.", window.TabDrafts.UpdateRequiredView!.Reason);
            Assert.Equal(1, asked());

            // Both covered: nothing more to ask.
            window.ReportApiOutdated(AppUpdateArea.Maintainer, "The server's own sentence.");
            await SettleAsync(() => true);

            Assert.Equal(1, asked());
        });
    }

    // ###########################################################################################
    // *** THE BUTTON IS THE WAY OUT. *** With an update already found it installs it; otherwise it
    // opens the download page - on either tab.
    // ###########################################################################################
    [Theory]
    [InlineData("3.1.0", true)]
    [InlineData(null, false)]
    public async Task The_button_installs_the_update_found_or_opens_the_download_page(string? pendingVersion, bool installs)
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + 1, pendingVersion);
            var pressed = new List<bool>();
            window.UpdateRequiredActionOverrideForTests = pressed.Add;

            await window.AskServerApiRevisionAsync();

            window.TabDrafts.GetControl<UpdateRequiredOverlay>("UpdateRequired").ClickActionForTests();
            window.TabMaintainer.GetControl<UpdateRequiredOverlay>("UpdateRequired").ClickActionForTests();

            Assert.Equal([installs, installs], pressed);
            Assert.Equal(
                installs ? AppUpdateRequiredWording.InstallButton : AppUpdateRequiredWording.DownloadPageButton,
                window.TabDrafts.UpdateRequiredView!.ButtonLabel);
        });
    }

    // An update found AFTER the tabs were covered changes their button (CheckForAppUpdateNowAsync
    // calls ApplyAppUpdateRequired once it has asked GitHub).
    [Fact]
    public async Task An_update_found_later_turns_the_button_into_Install_update()
    {
        await UiTest.RunAsync(async () =>
        {
            string? pending = null;
            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + 1);
            window.PendingVersionOverrideForTests = () => pending;

            await window.AskServerApiRevisionAsync();
            Assert.False(window.TabMaintainer.UpdateRequiredView!.ButtonInstalls);

            pending = "3.1.0";
            window.ApplyAppUpdateRequired();

            Assert.True(window.TabMaintainer.UpdateRequiredView!.ButtonInstalls);
            Assert.True(window.TabDrafts.UpdateRequiredView!.ButtonInstalls);
        });
    }

    // ###########################################################################################
    // *** A COVERED MAINTAINER TAB STOPS ASKING THE SERVER. *** Every minute check would be answered
    // "update CRT"; they stop when the tab is covered, and a remembered sign-in restored after that
    // starts none. The session itself is kept - a 426 is not a 401.
    // ###########################################################################################
    [Fact]
    public void A_covered_Maintainer_tab_stops_its_minute_checks_and_keeps_its_sign_in()
    {
        UiTest.Run(() =>
        {
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");
            AppUpdateRequiredView view = AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, "Please update CRT.", null);

            // Signed in first, then covered: the checks stop.
            var signedIn = MainUpdateRequiredTests.TabWithRememberedSession(session);
            signedIn.RestoreSignInQuietly();
            Assert.True(signedIn.QueueChecksRunningForTests);

            signedIn.ShowUpdateRequired(view);

            Assert.False(signedIn.QueueChecksRunningForTests);
            Assert.Same(session, signedIn.SignedIn);

            // Covered first, then the remembered sign-in restored: none start.
            var coveredFirst = MainUpdateRequiredTests.TabWithRememberedSession(session);
            coveredFirst.ShowUpdateRequired(view);
            coveredFirst.RestoreSignInQuietly();

            Assert.False(coveredFirst.QueueChecksRunningForTests);
            Assert.Same(session, coveredFirst.SignedIn);
        });
    }

    // ###########################################################################################
    // *** EDGE TO EDGE ON BOTH TABS. *** The Maintainer tab's root Grid has a 16px margin, and the
    // first render showed the overlay inside it - an uncovered frame round the scrim. Laid out in a
    // shown window, each tab's overlay starts at the tab's corner and is the tab's size.
    // ###########################################################################################
    [Fact]
    public void The_overlay_covers_each_tab_edge_to_edge()
    {
        UiTest.Run(() =>
        {
            AppUpdateRequiredView view = AppUpdateRequiredWording.For(AppUpdateArea.Drafts, "Please update CRT.", null);
            var drafts = new TabDrafts();
            var maintainer = new TabMaintainer();

            drafts.ShowUpdateRequired(view);
            maintainer.ShowUpdateRequired(view);

            foreach (UserControl tab in new UserControl[] { drafts, maintainer })
            {
                var window = new Window { Content = tab, Width = 900, Height = 600 };
                window.Show();
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();

                var overlay = tab.GetControl<UpdateRequiredOverlay>("UpdateRequired");

                Assert.Equal(new Point(0, 0), overlay.TranslatePoint(new Point(0, 0), tab)!.Value);
                Assert.Equal(tab.Bounds.Size, overlay.Bounds.Size);

                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** NOT ASKED FOR SOMEBODY WHO USES NEITHER TAB (code review, 2026-10-09). *** Every launch of
    // every installation sent the server a request it counts, for two tabs most never see. Now the
    // question goes out once the Drafts tab is shown or the Maintainer tab turned on - once, and
    // never before StartAsync lets it.
    // ###########################################################################################
    [Fact]
    public async Task The_server_is_asked_nothing_until_a_tab_it_is_about_is_in_use()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, asked) = BuildWindow(ClientVersionContract.ApiRevision);

            // Turned on before StartAsync: nothing yet.
            UserSettings.EnableMaintainerTab = true;
            window.ApplyMaintainerTabVisibility();
            UserSettings.EnableMaintainerTab = false;
            window.ApplyMaintainerTabVisibility();

            // No drafts, the Maintainer tab off: StartAsync's question does not go out.
            window.AllowApiRevisionQuestionForTests();
            await SettleAsync(() => true);
            Assert.Equal(0, asked());

            // Turned on now: asked - once, however often it is applied again.
            UserSettings.EnableMaintainerTab = true;
            window.ApplyMaintainerTabVisibility();
            await SettleAsync(() => asked() == 1);

            window.ApplyMaintainerTabVisibility();
            await SettleAsync(() => true);
            Assert.Equal(1, asked());
        });
    }

    // ###########################################################################################
    // *** INSTALLING THE UPDATE ASKS ABOUT UNSAVED TABLE EDITS FIRST (code review, 2026-10-09). ***
    // The restart exits CRT without the window's Closing event, so quitting's question was never
    // asked: the overlay's "Install update" threw away the edits in a Drafts table it had itself put
    // out of reach - under a line saying "nothing is lost". Cancel installs nothing; Save saves the
    // draft, then installs. The table's Save needs no server, so it works under the cover.
    // ###########################################################################################
    [Fact]
    public async Task Installing_the_update_asks_about_the_Drafts_tables_unsaved_edits_first()
    {
        await UiTest.RunAsync(async () =>
        {
            const string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
            CachedWorkbooks.Write(workbook, new BoardData { Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "PLA" }] });
            DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile), new DraftMarker { BoardKey = excelDataFile });

            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + 1, pendingVersion: "3.1.0");
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = excelDataFile };

            window.TabDrafts.HardwareBoardsOverrideForTests = [entry];
            window.TabDrafts.PublishedBoardOverrideForTests = _ => null;
            await window.TabDrafts.OpenTableAsync(entry);

            BoardTableSheet components = window.TabDrafts.TableEditor.SessionForTests!.Document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            components.Rows[0].Cells[components.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName)].Text = "Typed in the table";
            Assert.True(window.TabDrafts.HasUnsavedTableEdits);

            await window.AskServerApiRevisionAsync();
            Assert.True(window.TabDrafts.IsUpdateRequiredShown);

            var asked = new List<UnsavedTableEditsPrompt>();
            UnsavedTableEditsChoice answer = UnsavedTableEditsChoice.Cancel;
            window.TabDrafts.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return answer;
            };

            int installs = 0;
            window.DownloadAndInstallOverrideForTests = () =>
            {
                installs++;
                return Task.FromResult(false);
            };

            // Cancel: nothing installed, the edits kept.
            await window.InstallPendingUpdateAsync();

            Assert.Equal([UnsavedTableEditsPrompt.Leaving], asked);
            Assert.Equal(0, installs);
            Assert.True(window.TabDrafts.HasUnsavedTableEdits);

            // Save: into the draft, then installed.
            answer = UnsavedTableEditsChoice.Save;
            await window.InstallPendingUpdateAsync();

            Assert.Equal(1, installs);
            Assert.False(window.TabDrafts.HasUnsavedTableEdits);
            Assert.Equal("Typed in the table", DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, excelDataFile)!.Components[0].FriendlyName);
        });
    }

    // The Maintainer tab's tables too - asked without a Save while the tab is covered, since a save
    // would be turned away (LeavingUpdateRequired).
    [Fact]
    public async Task Installing_the_update_asks_about_a_Maintainer_tables_unsaved_change_first()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + 1, pendingVersion: "3.1.0");
            BoardDetailView view = window.TabMaintainer.BoardDetailForTests;
            var board = new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1);

            view.ShowDetailForTests(new BoardDetailAnswer(board, [], [], []));
            view.OpenTableForTests(new BoardTableAnswer(board.BoardId, "f", new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA" }] }, MayEdit: true));

            BoardTableSheet components = view.BoardTableForTests.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            components.Rows.Single().Cells[BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU 6510";

            await window.AskServerApiRevisionAsync();

            var asked = new List<UnsavedTableEditsPrompt>();
            UnsavedTableEditsChoice answer = UnsavedTableEditsChoice.Cancel;
            view.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return answer;
            };

            int installs = 0;
            window.DownloadAndInstallOverrideForTests = () =>
            {
                installs++;
                return Task.FromResult(false);
            };

            await window.InstallPendingUpdateAsync();

            Assert.Equal([UnsavedTableEditsPrompt.LeavingUpdateRequired], asked);
            Assert.Equal(0, installs);

            answer = UnsavedTableEditsChoice.Discard;
            await window.InstallPendingUpdateAsync();

            Assert.Equal(1, installs);
        });
    }

    // ###########################################################################################
    // *** NOTHING NEW IS MADE FOR A COVERED DRAFTS TAB (code review, 2026-10-09). *** "Edit board as
    // draft" and "Add a new board" exist to work in that tab; a draft made under its cover opened a
    // table nobody could reach. Both make nothing, and the Contribute tab says why.
    // ###########################################################################################
    [Fact]
    public async Task Edit_board_as_draft_and_Add_a_new_board_make_nothing_while_the_Drafts_tab_is_covered()
    {
        await UiTest.RunAsync(async () =>
        {
            const string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

            var (window, _) = BuildWindow(ClientVersionContract.ApiRevision + 1);
            window.EditAsDraftBoardOverrideForTests = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = excelDataFile };

            await window.AskServerApiRevisionAsync();
            Assert.True(window.TabDrafts.IsUpdateRequiredShown);

            TextBlock problem = window.TabContribute.GetControl<TextBlock>("DraftProblemText");

            await window.EditBoardAsDraftAsync();

            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, excelDataFile));
            Assert.True(problem.IsVisible);
            Assert.Equal(AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered, problem.Text);

            window.TabContribute.ShowDraftProblem(null);

            // No dialog is opened - one would block here - and the same line says why.
            await window.OpenNewBoardWindowAsync();

            Assert.True(problem.IsVisible);
            Assert.Equal(AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered, problem.Text);
        });
    }

    // ###########################################################################################
    // *** THE FIRST DRAFT OF A RUN ASKS THE SERVER FIRST (code review, 2026-10-10). *** With no
    // drafts and the Maintainer tab off, nothing had asked - so the guard above could not fire: the
    // draft was made, the Drafts tab appeared, its first showing asked, and the answer covered the
    // table just opened. Now "Edit board as draft" asks first and makes nothing; "Add a new board"
    // after it finds the tab covered without asking again - one question a run.
    // ###########################################################################################
    [Fact]
    public async Task The_first_draft_of_a_run_asks_the_server_first_and_makes_nothing_when_turned_away()
    {
        await UiTest.RunAsync(async () =>
        {
            const string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

            var (window, asked) = BuildWindow(ClientVersionContract.ApiRevision + 1);
            window.EditAsDraftBoardOverrideForTests = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = excelDataFile };

            // No drafts, the Maintainer tab off: StartAsync's question does not go out.
            window.AllowApiRevisionQuestionForTests();
            await SettleAsync(() => true);
            Assert.Equal(0, asked());
            Assert.False(window.TabDrafts.IsUpdateRequiredShown);

            TextBlock problem = window.TabContribute.GetControl<TextBlock>("DraftProblemText");

            await window.EditBoardAsDraftAsync();

            Assert.Equal(1, asked());
            Assert.True(window.TabDrafts.IsUpdateRequiredShown);
            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, excelDataFile));
            Assert.Equal(AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered, problem.Text);

            window.TabContribute.ShowDraftProblem(null);
            await window.OpenNewBoardWindowAsync();

            Assert.Equal(1, asked());
            Assert.Equal(AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered, problem.Text);
        });
    }

    // ###########################################################################################
    // *** ONE RECORD OF "SUBMISSIONS TURNED AWAY" (code review, 2026-10-10). *** The minute
    // submission check kept a flag of its own, set only by its own refusal: a 426 met by Submit, "My
    // submissions" or the discard notice covered the Drafts tab while the check went on asking every
    // minute - each refused, each asking /api/health again. Now whatever covers the Drafts tab stops
    // the check, and the check's own refusal covers the tab. A Maintainer-only refusal stops nothing.
    // ###########################################################################################
    [Fact]
    public async Task Whatever_turns_submissions_away_covers_the_Drafts_tab_and_stops_the_minute_check()
    {
        await UiTest.RunAsync(async () =>
        {
            const string Words = "Please update CRT.";

            // A refusal met elsewhere - Submit, say.
            var (refused, refusedAsked) = BuildWindow(ClientVersionContract.ApiRevision);
            Assert.False(refused.SubmissionChecksOutdatedForTests);

            refused.ReportApiOutdated(AppUpdateArea.Drafts, Words);
            await SettleAsync(() => refusedAsked() == 1);

            Assert.True(refused.TabDrafts.IsUpdateRequiredShown);
            Assert.True(refused.SubmissionChecksOutdatedForTests);

            // The server on a higher API revision.
            var (behind, _) = BuildWindow(ClientVersionContract.ApiRevision + 1);
            await behind.AskServerApiRevisionAsync();

            Assert.True(behind.SubmissionChecksOutdatedForTests);

            // The minute check's own refusal: the Drafts tab is covered too.
            var (statusCheck, _) = BuildWindow(ClientVersionContract.ApiRevision);
            statusCheck.ShowSubmissionChecksOutdated(Words);

            Assert.True(statusCheck.SubmissionChecksOutdatedForTests);
            Assert.True(statusCheck.TabDrafts.IsUpdateRequiredShown);
            Assert.Equal(Words, statusCheck.TabDrafts.UpdateRequiredView!.Reason);

            // A Maintainer-only refusal leaves submissions alone.
            var (maintainerOnly, maintainerAsked) = BuildWindow(ClientVersionContract.ApiRevision);
            maintainerOnly.ReportApiOutdated(AppUpdateArea.Maintainer, Words);
            await SettleAsync(() => maintainerAsked() == 1);

            Assert.False(maintainerOnly.SubmissionChecksOutdatedForTests);
        });
    }

    // ###########################################################################################
    // *** THE COVER IS THE ONE RECORD (code review, 2026-10-10). *** The Maintainer tab kept a flag
    // beside its overlay, set only by ShowUpdateRequired - the overlay shown any other way left the
    // minute checks free to start. Now they read the overlay itself.
    // ###########################################################################################
    [Fact]
    public void However_the_Maintainer_tab_is_covered_no_minute_check_starts()
    {
        UiTest.Run(() =>
        {
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");
            var tab = MainUpdateRequiredTests.TabWithRememberedSession(session);

            tab.UpdateRequired.Show(AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, "Please update CRT.", null));
            tab.RestoreSignInQuietly();

            Assert.True(tab.IsUpdateRequiredShown);
            Assert.False(tab.QueueChecksRunningForTests);
            Assert.Same(session, tab.SignedIn);
        });
    }

    // A tab whose remembered sign-in and server are the test's own - never the user's file or the
    // real server.
    private static TabMaintainer TabWithRememberedSession(ReviewSession session)
    {
        var tab = new TabMaintainer
        {
            RecallSessionOverrideForTests = _ => session,
            ClientOverrideForTests = () => new ReviewApiClient(
                "https://review.invalid",
                new HttpClient(new ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler(
                    _ => ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler.Refused())))
        };

        return tab;
    }

    // Lets the dispatcher run what a fire-and-forget question queued, until `done` holds - never a
    // fixed sleep. Fails rather than hanging if it never does.
    private static async Task SettleAsync(Func<bool> done)
    {
        for (int attempt = 0; attempt < 500 && !done(); attempt++)
        {
            await Task.Delay(2);
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        }

        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        Assert.True(done());
    }
}
