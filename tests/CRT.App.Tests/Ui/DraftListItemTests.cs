using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// One row of the Drafts tab - what it says and what Submit allows (code review, 2026-10-13).
//
// Built directly: the row decides from what it is handed, and these pin those decisions.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftListItemTests
{
    private static readonly HardwareBoardEntry Entry = new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        ExcelDataFile = "Commodore/C64/250407/Data C64 250407.xlsx",
    };

    private static DraftListItem Row(
        int rows = 0,
        DraftOtherChanges other = default,
        DraftDriftState drift = DraftDriftState.InSync,
        bool hasBase = true,
        SubmissionReceipt? lastSubmission = null,
        bool alreadySent = false) =>
        new(
            Entry,
            new DraftStatus { BoardKey = Entry.ExcelDataFile, BaseRevision = "2026-09-01", HasBase = hasBase },
            rows,
            BoardProblemCounts.None,
            drift,
            _ => Task.CompletedTask,
            _ => Task.CompletedTask,
            _ => Task.CompletedTask,
            isTableOpen: false,
            lastSubmission,
            alreadySent,
            other);

    // The bug: a draft whose only work is a KiCad calibration or a picture could not be sent.
    [Fact]
    public void A_draft_holding_only_a_calibration_or_files_can_be_submitted_and_says_what_it_holds()
    {
        UiTest.Run(() =>
        {
            DraftListItem calibration = Row(other: new DraftOtherChanges(0, 1));

            Assert.True(calibration.CanSubmit);
            Assert.Equal("KiCad calibration changed", calibration.RowSummary);
            Assert.False(calibration.HasNote);

            DraftListItem both = Row(rows: 2, other: new DraftOtherChanges(1, 0));

            Assert.True(both.CanSubmit);
            Assert.Equal("2 rows changed, 1 file changed", both.RowSummary);
        });
    }

    [Fact]
    public void An_unchanged_draft_cannot_be_sent_and_the_row_says_why()
    {
        UiTest.Run(() =>
        {
            DraftListItem row = Row();

            Assert.False(row.CanSubmit);
            Assert.True(row.HasNote);
            Assert.Equal(row.SubmitTooltip, row.NoteText);
            Assert.Contains("no change from the published data", row.NoteText);
        });
    }

    // Why Submit is grey is said ON the row, not only in a tooltip a keyboard user never sees.
    [Fact]
    public void A_draft_already_sent_as_it_is_says_so_on_the_row()
    {
        UiTest.Run(() =>
        {
            var sent = new SubmissionReceipt
            {
                SubmissionId = 12,
                BoardId = "Commodore/C64/250407",
                SentUtc = new DateTimeOffset(2026, 10, 4, 12, 0, 0, TimeSpan.Zero),
                LastKnownState = "pending",
                DraftFingerprint = "v1:ABC",
            };

            DraftListItem row = Row(rows: 1, lastSubmission: sent, alreadySent: true);

            Assert.False(row.CanSubmit);
            Assert.Contains("already sent this draft", row.NoteText);
        });
    }

    [Fact]
    public void A_drifted_draft_with_a_base_is_told_to_bring_it_up_to_date()
    {
        UiTest.Run(() =>
        {
            DraftListItem row = Row(rows: 1, drift: DraftDriftState.OfficialIsNewer, hasBase: true);

            Assert.True(row.HasDrift);
            Assert.Contains("Bring your draft up to date", row.DriftSummary);
            Assert.DoesNotContain("still applied", row.DriftSummary);
        });
    }

    // ###########################################################################################
    // A submission a maintainer corrected and published leaves its draft behind (retirement keeps
    // anything that differs from the published board), and sending it again would undo the
    // correction. The row says the draft is finished with, and shows no drift - the "update" is
    // the contributor's own work as published.
    // ###########################################################################################
    [Fact]
    public void A_published_and_amended_submission_says_the_draft_is_finished_with()
    {
        UiTest.Run(() =>
        {
            var published = new SubmissionReceipt
            {
                SubmissionId = 12,
                BoardId = "Commodore/C64/250407",
                SentUtc = new DateTimeOffset(2026, 10, 4, 12, 0, 0, TimeSpan.Zero),
                LastKnownState = "published",
                AmendedByMaintainer = true,
                DraftFingerprint = "v1:ABC",
            };

            DraftListItem row = Row(rows: 1, drift: DraftDriftState.OfficialIsNewer, lastSubmission: published, alreadySent: true);

            Assert.False(row.CanSubmit);
            Assert.False(row.HasDrift);
            Assert.Contains("discard this draft", row.NoteText);
            Assert.Contains("undo the maintainer's changes", row.NoteText);
        });
    }
}
