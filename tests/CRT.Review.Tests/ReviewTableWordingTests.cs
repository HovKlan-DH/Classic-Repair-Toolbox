using System.Text.Json;
using CRT.Review.Handlers;
using Handlers.DataHandling;

namespace CRT.Review.Tests;

// Covers ReviewTableWording and the parsing behind the reviewer's table (maintainer request,
// 2026-09-25: "View in table format" in the review application, where a reviewer may also edit).
//
// The one decision here is RowsToSave - what a save sends. The rest is wording, tested so it
// stays true to what the server does with an amendment.
public sealed class ReviewTableWordingTests
{
    private static SubmissionRows Submitted() => new()
    {
        RevisionDate = "2026-September-25",
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
        ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
        KiCadCalibrations = [new KiCadCalibrationEntry()]
    };

    // ---- what a save sends ---------------------------------------------------------------

    [Fact]
    public void A_save_sends_the_submission_with_the_tables_edits_applied_and_nothing_else_lost()
    {
        SubmissionRows submitted = ReviewTableWordingTests.Submitted();
        BoardTableDocument document = BoardTableDocument.Create(null, SubmissionRowsBoard.ToBoard(submitted));

        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        components.Rows.Single().Cells[friendly].Text = "CPU 6510";

        SubmissionRows sent = ReviewTableWording.RowsToSave(document, submitted);

        Assert.Equal("CPU 6510", Assert.Single(sent.Components).FriendlyName);

        // What the table does not show travels unchanged - the server keeps it regardless, but a
        // request that dropped it would still read as a deletion to anyone looking at it.
        Assert.Equal("2026-September-25", sent.RevisionDate);
        Assert.Single(sent.ComponentHighlights);
        Assert.Single(sent.KiCadCalibrations);
    }

    // ###########################################################################################
    // *** A REVIEWER DELETING A COMPONENT DELETES ITS HIGHLIGHTS TOO (2026-09-25) - BOTH ENDS. ***
    // The table drops them from what it sends; the server keeps a highlight unless the edit dropped
    // it AND no component in the edit has its label. Run through both here, the review app's save
    // and the server's rule, so the two cannot disagree about it.
    // ###########################################################################################
    [Fact]
    public void Deleting_a_component_in_the_table_removes_its_highlights_from_the_amendment()
    {
        SubmissionRows submitted = ReviewTableWordingTests.Submitted();
        submitted.Components.Add(new ComponentEntry { BoardLabel = "U9", FriendlyName = "CIA", TechnicalNameOrValue = "6526" });
        submitted.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U9", X = "5", Y = "5", Width = "3", Height = "4" });

        BoardTableDocument document = BoardTableDocument.Create(null, SubmissionRowsBoard.ToBoard(submitted));
        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int label = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);
        components.DeleteRow(components.Rows.First(row => row.Cells[label].Text == "U8"));

        SubmissionRows sent = ReviewTableWording.RowsToSave(document, submitted);
        SubmissionRows amended = SubmissionRowsBoard.WithTableSections(submitted, sent);

        Assert.Equal("U9", Assert.Single(sent.ComponentHighlights).BoardLabel);
        Assert.Equal("U9", Assert.Single(amended.ComponentHighlights).BoardLabel);
    }

    // ---- the words -----------------------------------------------------------------------

    [Fact]
    public void The_table_says_what_it_is_coloured_against()
    {
        Assert.StartsWith("Nothing of this system is published yet",
            ReviewTableWording.OpenedMessage(new ReviewTableData(0, null, ReviewTableWordingTests.Submitted())));

        Assert.StartsWith("Coloured against the published board",
            ReviewTableWording.OpenedMessage(new ReviewTableData(0, ReviewTableWordingTests.Submitted(), ReviewTableWordingTests.Submitted())));
    }

    // Saving clears approvals given earlier - said, so a reviewer who had approved knows to again.
    [Fact]
    public void A_save_says_approvals_were_cleared_and_repeats_any_warning()
    {
        Assert.Contains("approval given before this change was cleared",
            ReviewTableWording.Saved(new ReviewAmendResult(1, [])), StringComparison.Ordinal);

        string withWarning = ReviewTableWording.Saved(new ReviewAmendResult(2,
            [new ReviewFindingView("row.x", "U8", "U8 has no category.", IsError: false)]));

        Assert.EndsWith("Warnings: U8 has no category.", withWarning, StringComparison.Ordinal);
    }

    [Fact]
    public void A_refused_save_says_why()
    {
        Assert.Equal("Not saved: Anna changed this submission after you opened it.",
            ReviewTableWording.NotSaved("Anna changed this submission after you opened it."));
        Assert.Equal("Not saved: the server refused it.", ReviewTableWording.NotSaved(""));
    }

    [Fact]
    public void The_submission_view_names_who_changed_it()
    {
        Assert.Null(ReviewTableWording.AmendedLine(null));

        string? line = ReviewTableWording.AmendedLine(new ReviewAmendmentView(1, "Anna (anna@example.com)", null));

        Assert.StartsWith("Changed in the table by Anna (anna@example.com).", line, StringComparison.Ordinal);
        Assert.Contains("the contributor is told", line, StringComparison.Ordinal);
    }

    [Fact]
    public void The_window_names_the_submission_and_its_board()
    {
        var row = new ReviewQueueRow(42, "Commodore/C64/250407", "pending", "x", "c@example.com", null);

        Assert.Equal("Submission #42 - Commodore/C64/250407 - table", ReviewTableWording.WindowTitle(row));
        Assert.Equal("View in table format", ReviewTableWording.ButtonText);
    }

    // ---- reading it off the server -------------------------------------------------------

    [Fact]
    public void The_table_is_read_as_the_servers_own_record()
    {
        string json = JsonSerializer.Serialize(
            new ReviewTableData(3, null, ReviewTableWordingTests.Submitted()), JsonSerializerOptions.Web);

        ReviewTableData? table = ReviewApiParser.ParseTable(json);

        Assert.Equal(3, table!.Version);
        Assert.Null(table.Published);
        Assert.Equal("U8", Assert.Single(table.Submitted.Components).BoardLabel);
    }

    [Fact]
    public void A_saved_change_is_read_with_its_version_and_warnings()
    {
        ReviewAmendResult? result = ReviewApiParser.ParseAmend(
            """{"version":2,"findings":[{"severity":"Warning","code":"row.x","subject":"U8","message":"U8 has no category."}]}""");

        Assert.Equal(2, result!.Version);
        Assert.Equal("U8 has no category.", Assert.Single(result.Warnings).Message);
        Assert.Null(ReviewApiParser.ParseAmend("""{"findings":[]}"""));
    }

    [Fact]
    public void A_detail_names_who_last_changed_the_submission()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"canPublish":true,
             "submission":{"id":42,"systemId":"Commodore/C64/250407","state":"pending","summary":"x"},
             "findings":[],
             "amendment":{"version":2,"by":"Anna (anna@example.com)","atUtc":"2026-09-25T12:00:00+00:00"}}
            """);

        Assert.Equal(2, detail!.Amendment!.Version);
        Assert.Equal("Anna (anna@example.com)", detail.Amendment.By);
        Assert.NotNull(detail.Amendment.AtUtc);
    }
}
