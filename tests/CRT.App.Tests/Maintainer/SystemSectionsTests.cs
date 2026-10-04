using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers SystemSections - a system's six views on the Systems screen (owner request, 2026-10-03:
// "Board data ... Files ... Contributor ... Maintainer ... Statistics", and 2026-10-04: "another
// 'History' tab/button, after the 'Maintainer' button"), and what its Board data
// view says and sends. A save there goes STRAIGHT TO BETA (owner decision, 2026-10-03: "it should go
// directly to the next queue, 'BETA > Stable'"), after a reason is asked for - so the words say so
// before it happens, and say plainly afterwards whether it was published or only made into a
// submission that waits.
// ###########################################################################################
public sealed class SystemSectionsTests
{
    private const string SystemId = "Commodore/C64/250407";

    [Fact]
    public void The_six_views_are_in_the_order_the_project_owner_named_them_and_open_on_Board_data()
    {
        Assert.Equal(
            ["Board data", "Files", "Contributor", "Maintainer", "History", "Statistics"],
            SystemSections.Order.Select(SystemSections.Label));

        Assert.Equal(SystemSection.BoardData, SystemSections.Opening);
        Assert.Equal(SystemSections.Opening, SystemSections.Order[0]);

        // The same word as a submission's own Board data button.
        Assert.Equal(SubmissionViews.Label(SubmissionView.BoardData), SystemSections.Label(SystemSection.BoardData));
    }

    // ###########################################################################################
    // What the table says when it opens: for a system this account may change, that a save asks for
    // a reason and goes straight to BETA, then waits under "Queue: Awaiting push from BETA to stable" - no longer that it becomes
    // a submission to approve; for one it may not, the server's reason.
    // ###########################################################################################
    [Fact]
    public void An_editable_table_says_a_save_asks_a_reason_and_goes_straight_to_BETA()
    {
        string opened = SystemSections.OpenedMessage(new SystemTableAnswer(SystemSectionsTests.SystemId, "f", new SubmissionRows(), MayEdit: true));

        Assert.Contains("asks for a reason", opened, StringComparison.Ordinal);
        Assert.Contains("straight to BETA", opened, StringComparison.Ordinal);
        Assert.Contains("waits under \"Queue: Awaiting push from BETA to stable\"", opened, StringComparison.Ordinal);
        Assert.DoesNotContain("Contributor submissions", opened, StringComparison.Ordinal);
    }

    [Fact]
    public void A_table_this_account_may_not_change_says_why_in_the_servers_words()
    {
        Assert.Equal(
            "Only this system's maintainers can.",
            SystemSections.OpenedMessage(new SystemTableAnswer(SystemSectionsTests.SystemId, "f", new SubmissionRows(), MayEdit: false, "Only this system's maintainers can.")));

        // No reason from the server: still said that it cannot be changed.
        Assert.Contains(
            "not send a change",
            SystemSections.OpenedMessage(new SystemTableAnswer(SystemSectionsTests.SystemId, "f", new SubmissionRows(), MayEdit: false)),
            StringComparison.Ordinal);
    }

    // ###########################################################################################
    // THE BETA / STABLE SWITCH (owner request, 2026-10-04): the half wanted when the system is there;
    // the stable source when only it holds the system; else BETA. A half the system is not in cannot
    // be chosen - except BETA when neither holds it, whose view then says so.
    // ###########################################################################################
    [Theory]
    [InlineData(false, true, true, false)]
    [InlineData(true, true, true, true)]
    [InlineData(true, true, false, false)]
    [InlineData(true, true, null, false)]
    [InlineData(false, false, true, true)]
    [InlineData(true, false, false, false)]
    public void The_switch_shows_the_half_wanted_where_the_system_is(bool stableWanted, bool inBeta, bool? inStable, bool showsStable)
    {
        var system = new SystemOverviewEntry(SystemSectionsTests.SystemId, "Commodore", "C64", "250407", inBeta, inStable, false, true, null, null, null, 1);

        Assert.Equal(showsStable, SystemSections.ShowsStable(stableWanted, system));
        Assert.False(SystemSections.ShowsStable(true, null));
    }

    [Theory]
    [InlineData(true, true, true, true)]
    [InlineData(true, false, true, false)]
    [InlineData(false, true, false, true)]
    [InlineData(false, false, true, false)]
    [InlineData(false, null, true, false)]
    public void A_half_the_system_is_not_in_cannot_be_chosen(bool inBeta, bool? inStable, bool canBeta, bool canStable)
    {
        var system = new SystemOverviewEntry(SystemSectionsTests.SystemId, "Commodore", "C64", "250407", inBeta, inStable, false, true, null, null, null, 1);

        Assert.Equal(canBeta, SystemSections.CanChooseBeta(system));
        Assert.Equal(canStable, SystemSections.CanChooseStable(system));
    }

    // An older server ignores the request's tree and answers with BETA's board and files - which
    // must never be drawn under "Stable" (2026-10-04).
    [Fact]
    public void Only_an_answer_from_the_stable_source_counts_as_stable()
    {
        var rows = new SubmissionRows();

        Assert.True(SystemSections.IsStableAnswer(new SystemTableAnswer(SystemSectionsTests.SystemId, "f", rows, MayEdit: false, "x", BetaDataUrl: null, ProductionDataUrl: "https://example.org/")));
        Assert.False(SystemSections.IsStableAnswer(new SystemTableAnswer(SystemSectionsTests.SystemId, "f", rows, MayEdit: true, BetaDataUrl: "https://example.org/beta")));

        Assert.True(SystemSections.IsStableAnswer(new SystemFilesAnswer(SystemSectionsTests.SystemId,
            [new SystemFileEntry("a/b.png", SystemFileChange.Unchanged, SystemFileSource.Production)], null, "https://example.org/")));
        Assert.False(SystemSections.IsStableAnswer(new SystemFilesAnswer(SystemSectionsTests.SystemId,
            [new SystemFileEntry("a/b.png", SystemFileChange.Unchanged, SystemFileSource.Beta)], "https://example.org/beta")));
    }

    // ###########################################################################################
    // The reason dialog's words: what happens on Publish (BETA now, then "Queue: Awaiting push from BETA to stable", nothing else
    // meanwhile), how to try it - naming the Configuration tab's check box by its own constant - and
    // a button that says what it does, with no "..." (the button rule).
    // ###########################################################################################
    [Fact]
    public void The_reason_dialog_says_where_the_change_goes_and_how_to_try_it()
    {
        Assert.Contains("straight into the BETA data", SystemSections.ReasonExplanation, StringComparison.Ordinal);
        Assert.Contains(ConfigurationWording.BetaSourceCheckBox, SystemSections.ReasonExplanation, StringComparison.Ordinal);
        Assert.Contains("waits under \"Queue: Awaiting push from BETA to stable\"", SystemSections.ReasonExplanation, StringComparison.Ordinal);
        Assert.Contains("no other change", SystemSections.ReasonExplanation, StringComparison.Ordinal);

        Assert.Equal("Publish to BETA", SystemSections.PublishButton);
        Assert.Equal("Reason for the change", SystemSections.ReasonLabel);
        Assert.Equal("Publish your change to Commodore/C64/250407 in BETA", SystemSections.ReasonHeadline(SystemSectionsTests.SystemId));
    }

    // Files the publish removes are named before it - none, no heading at all; one and several counted.
    [Fact]
    public void The_files_a_change_removes_get_a_counted_heading_and_none_get_none()
    {
        Assert.Null(SystemSections.RemovalsHeading([]));
        Assert.Equal(
            "This also removes 1 file from BETA - nothing uses it once your change is in:",
            SystemSections.RemovalsHeading(["Commodore/C64/250407/manual.pdf"]));
        Assert.Equal(
            "This also removes 2 files from BETA - nothing uses them once your change is in:",
            SystemSections.RemovalsHeading(["Commodore/C64/250407/a.pdf", "Commodore/C64/250407/b.pdf"]));
    }

    [Fact]
    public void A_published_change_names_its_revision_what_it_removed_and_any_warnings()
    {
        Assert.Equal(
            "Published to BETA as revision 2026-October-03.",
            SystemSections.Published(new SystemEditResult(57, [], Published: true, Revision: "2026-October-03", RemovedFiles: [])));

        Assert.Equal(
            "Published to BETA as revision 2026-October-03. 1 file nothing used any more was removed.",
            SystemSections.Published(new SystemEditResult(57, [], Published: true, Revision: "2026-October-03", RemovedFiles: ["a.pdf"])));

        Assert.Equal(
            "Published to BETA. 2 files nothing used any more were removed. Warnings: Check U8.",
            SystemSections.Published(new SystemEditResult(
                57, [new ReviewFindingView("x", "", "Check U8.", IsError: false)], Published: true, RemovedFiles: ["a.pdf", "b.pdf"])));
    }

    // ###########################################################################################
    // Made into a submission but NOT published: said as plainly - never "published" - with the
    // server's reason and where it waits, so the maintainer knows nothing is lost and nothing is in
    // BETA yet.
    // ###########################################################################################
    [Fact]
    public void A_change_saved_but_not_published_says_so_why_and_where_it_waits()
    {
        string line = SystemSections.SavedNotPublished(new SystemEditResult(57, [], NotPublishedReason: "BETA moved."));

        Assert.Equal(
            "Saved as submission #57, but not published to BETA: BETA moved. It waits under \"Queue: Contributor submissions\", where you can approve it.",
            line);

        Assert.Contains("the server did not say why.", SystemSections.SavedNotPublished(new SystemEditResult(57, [])), StringComparison.Ordinal);
    }

    [Fact]
    public void A_refused_change_says_it_was_not_published_and_why()
    {
        Assert.Equal("Not published: BETA changed.", SystemSections.NotPublished("BETA changed."));
        Assert.Equal("Not published: the server refused it.", SystemSections.NotPublished(" "));
    }

    // ###########################################################################################
    // AFTER NO ANSWER IN TWO MINUTES the system's submissions say what happened: the NEWEST one with
    // the reason given (a reason used before must not be taken for this one), and whether it reached
    // BETA - "merged", or "published" once it went on to stable - or waits in the queue.
    // ###########################################################################################
    [Fact]
    public void After_a_timeout_the_newest_submission_with_the_reason_is_the_change_and_its_state_says_how_far_it_got()
    {
        static SystemSubmissionEntry Entry(long id, string summary, string state) =>
            new(id, null, summary, state, DateTimeOffset.UtcNow, null, null);

        SystemOverviewEntry system = new(SystemSectionsTests.SystemId, "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1);

        var detail = new SystemDetailAnswer(
            system, [], [],
            [Entry(40, "Corrected U8.", "published"), Entry(57, " Corrected U8. ", "merged"), Entry(58, "Something else.", "pending")]);

        SystemSubmissionEntry found = SystemSections.FindSent(detail, "Corrected U8.")!;

        Assert.Equal(57, found.Id);
        Assert.True(SystemSections.ReachedBeta(found));
        Assert.True(SystemSections.ReachedBeta(Entry(1, "x", "published")));
        Assert.False(SystemSections.ReachedBeta(Entry(1, "x", "pending")));
        Assert.Null(SystemSections.FindSent(detail, "Never sent."));
    }

    // ###########################################################################################
    // WHAT A SAVE SENDS: BETA's rows with the edit applied, and what the table cannot show - the
    // revision date, the highlights, the KiCad calibrations - as BETA has them. Dropped, approving the
    // submission would publish the board without its calibration work.
    // ###########################################################################################
    [Fact]
    public void A_save_sends_BETAs_rows_with_the_edit_and_keeps_what_the_table_cannot_show()
    {
        var beta = new SubmissionRows
        {
            RevisionDate = "2026-August-21",
            Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
            KiCadCalibrations = [new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "board" }]
        };

        BoardTableDocument document = BoardTableDocument.Create(SubmissionRowsBoard.ToBoard(beta), SubmissionRowsBoard.ToBoard(beta));

        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        components.Rows.Single().Cells[friendly].Text = "CPU 6510";

        SubmissionRows sent = SystemSections.RowsToSend(document, beta);

        Assert.Equal("CPU 6510", Assert.Single(sent.Components).FriendlyName);
        Assert.Equal("2026-August-21", sent.RevisionDate);
        Assert.Single(sent.ComponentHighlights);
        Assert.Equal("board", Assert.Single(sent.KiCadCalibrations).CadName);
    }

    // ###########################################################################################
    // A system waiting for a place in CRT's drop-down lists is said ABOVE the views, with where the
    // place is given - the view chosen for another system stays chosen, so the Maintainer view's
    // placement could otherwise go unseen. Placed already: the list's own mark alone.
    // ###########################################################################################
    [Fact]
    public void A_system_needing_a_place_says_so_and_where_and_a_placed_one_only_says_so()
    {
        var suggested = new SystemPlacement("C64", "250407", string.Empty, null);
        var unlisted = new UnlistedSystemEntry(SystemSectionsTests.SystemId, "Commodore", "C64", "250407", InBeta: true, CanPlace: true, Placement: null, Suggested: suggested);

        Assert.Equal(
            "Needs a place in the drop-down lists - give it one under Maintainer",
            SystemSections.PlacementLine(new SystemListingAnswer(true, [], [unlisted]), SystemSectionsTests.SystemId));

        Assert.Equal(
            "Placed - listed when published to BETA",
            SystemSections.PlacementLine(new SystemListingAnswer(true, [], [unlisted with { Placement = suggested }]), SystemSectionsTests.SystemId));

        Assert.Null(SystemSections.PlacementLine(new SystemListingAnswer(true, [], []), SystemSectionsTests.SystemId));
        Assert.Null(SystemSections.PlacementLine(null, SystemSectionsTests.SystemId));
    }
}
