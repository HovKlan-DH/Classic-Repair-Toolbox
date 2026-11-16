using Avalonia.Controls;
using Avalonia.Headless;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// What the Submit dialog sends, now that it asks for the address alone (owner requests, 2026-11-15: the credit "remove it from
// this modal, and have this directly done in the "Board: Credits" data", and of "What did you change?": "so remove it"). What the
// submission says it changed is handed in, made by the Contribute tab of the credit sent with the work (SubmissionSummaryText) -
// the dialog sends it as it is. ONE board per submit, the same day.
//
// COLLECTION NOTE: "HeadlessUi" because it builds a Window; it points UserSettings at a temp file first (the dialog reads and
// remembers the address there), as MainWindowTests does.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowSendingTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SubmitDraftWindowSendingTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private SubmitDraftWindow BuildWindow(BoardData board, string summary = "")
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Commodore 64 - 250407",
            board,
            new SubmissionIdentity { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" },
            this.thisWorkspace.Path_("draft"),
            this.thisWorkspace.Path_("Data"),
            draftFingerprint: "before",
            summary: summary);

        window.GetControl<TextBox>("EmailTextBox").Text = "bo@example.com";
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        return window;
    }

    // Nothing to type but the address: no credit, no name, no "What did you change?" - and Submit waits for nothing else.
    [Fact]
    public void The_dialog_asks_for_the_address_alone()
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData());

            foreach (string gone in new[] { "SummaryTextBox", "CreditOfferPanel", "NoCreditCheckBox", "CreditNameTextBox", "CreditContributionTextBox" })
            {
                Assert.Null(window.FindControl<Control>(gone));
            }

            Assert.True(window.GetControl<Button>("SubmitButton").IsEnabled);
        });
    }

    // What the maintainers' queue, their mail and "My submissions" show is the summary handed in, as it is.
    [Fact]
    public void The_submission_says_the_summary_it_was_handed()
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), summary: "Renamed U8 to its part number");

            window.CaptureEntriesForTests();

            Assert.Equal("Renamed U8 to its part number", window.NewReceipt(1, "token", DateTimeOffset.UtcNow).Summary);
        });
    }

    // The board is sent as handed in - its credit a row of it - and what came of it said, the dialog's Result the server's answer.
    [Fact]
    public async Task The_board_is_sent_as_it_was_handed_in_and_its_outcome_said()
    {
        var board = new BoardData { Credits = [new CreditEntry { Category = "Board data", NameOrHandle = "Bo", MainContribution = "Drew the bottom side" }] };
        var sent = new List<SubmitDraftTarget>();

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(board);

            window.SendOverrideForTests = (target, _) =>
            {
                sent.Add(target);
                return Task.FromResult(new SubmissionResult { SubmissionId = 42, IsAccepted = true, State = "pending" });
            };

            await window.SubmitNowAsync();

            SubmitDraftTarget target = Assert.Single(sent);
            Assert.Same(board, target.Board);
            Assert.Equal("before", target.Fingerprint);

            Assert.Equal(42L, window.Result?.SubmissionId ?? 0L);
            Assert.Equal("Contribution sent", window.GetControl<TextBlock>("HeaderText").Text);
            Assert.Contains("Your draft is untouched", window.GetControl<TextBlock>("OutcomeText").Text ?? string.Empty, StringComparison.Ordinal);
        });
    }
}
