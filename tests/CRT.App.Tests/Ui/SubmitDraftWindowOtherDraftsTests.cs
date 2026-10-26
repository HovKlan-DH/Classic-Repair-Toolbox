using System.Net.Http;
using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog's other drafts (CRT 4.0 step 2, slice 5 - the plan: "Submit offers to send all
// the drafts at once, with one message"). Listed unticked, greyed out with the reason when one cannot
// go; each ticked one sent AFTER this board's as a submission of its own, with the same message - made
// ready by the Contribute tab only once Submit is pressed, under the credit name typed here. What
// came of each board is listed by name; one that could not go never stops the rest.
//
// Nothing is sent over the network (test rule 6): SendOverrideForTests stands in for the send.
//
// COLLECTION NOTE: "HeadlessUi" because it builds a Window; UserSettings points at a temp file (the
// dialog reads and remembers the credit name there).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowOtherDraftsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SubmitDraftWindowOtherDraftsTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string board) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = board,
        ExcelDataFile = $"Commodore/C64/{board}/board.json",
    };

    private static readonly OtherDraftOffer[] Offers =
    [
        new(Entry("250425"), "Commodore 64 - 250425", "12 rows changed", null),
        new(Entry("250466"), "Commodore 64 - 250466", "1 row changed", "already sent as it is"),
        new(Entry("250469"), "Commodore 64 - 250469", "3 rows changed", null),
    ];

    private SubmitDraftWindow BuildWindow(IReadOnlyList<CreditSuggestion>? ownCredits = null)
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Commodore 64 - 250407",
            new BoardData(),
            new SubmissionIdentity { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" },
            this.thisWorkspace.Path_("draft"),
            this.thisWorkspace.Path_("Data"),
            draftFingerprint: "before",
            creditSuggestions: ownCredits ?? [],
            addCredits: _ => Task.FromResult<(BoardData, string)?>(null));

        window.EmailTextBox.Text = "bo@example.org";
        window.SummaryTextBox.Text = "Corrected the VIC's value";
        return window;
    }

    private static SubmitDraftTarget Target(string displayName, bool credited = false) => new(
        displayName,
        new BoardData(),
        new SubmissionIdentity { BoardId = "Commodore/C64/x", Manufacturer = "Commodore", Hardware = "C64", Board = "x" },
        string.Empty,
        [],
        "fingerprint",
        credited);

    private static SubmissionResult Accepted(long id) => new() { SubmissionId = id, IsAccepted = true, State = "pending" };

    [Fact]
    public void The_other_drafts_are_offered_unticked_and_one_that_cannot_go_says_why()
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            Assert.Empty(window.OtherDraftChoicesForTests);

            window.OfferOtherDrafts(Offers, (_, _, _) => Task.FromResult(OtherDraftReady.NotSent("x")));

            Assert.Equal(
                [
                    ("Commodore 64 - 250425 (12 rows changed)", true),
                    ("Commodore 64 - 250466 - already sent as it is", false),
                    ("Commodore 64 - 250469 (3 rows changed)", true),
                ],
                window.OtherDraftChoicesForTests);

            Assert.All(window.OtherDraftsPanel.Children.OfType<CheckBox>(), box => Assert.False(box.IsChecked));
            Assert.Equal(SubmitDraftsWording.OtherDraftsHeading, window.OtherDraftsHeadingText.Text);
        });
    }

    // With no credit of its own to offer, the name is still asked for once another draft is ticked -
    // that draft is credited under it.
    [Fact]
    public void Ticking_another_draft_asks_for_the_credit_name_it_goes_under()
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.OfferOtherDrafts(Offers, (_, _, _) => Task.FromResult(OtherDraftReady.NotSent("x")));
            Assert.False(window.CreditOfferPanel.IsVisible);

            window.TickOtherDraftForTests(0);

            Assert.True(window.CreditOfferPanel.IsVisible);
            Assert.False(window.CreditChoicesPanel.IsVisible);
            Assert.True(window.CreditOthersCheckBox.IsVisible);
            Assert.True(window.CreditOthersCheckBox.IsChecked);

            window.TickOtherDraftForTests(0, ticked: false);
            Assert.False(window.CreditOfferPanel.IsVisible);
        });
    }

    // ###########################################################################################
    // This board first, then each ticked draft made ready under the credit typed and sent - one that
    // cannot be made ready, or that never reaches the server, said by name while the rest go.
    // ###########################################################################################
    [Fact]
    public async Task Each_ticked_draft_is_sent_after_this_one_and_what_came_of_each_is_listed()
    {
        var sent = new List<string>();
        var prepared = new List<(string Board, string Name, string Contact)>();

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();

            window.OfferOtherDrafts(
                [.. Offers, new OtherDraftOffer(Entry("250470"), "Commodore 64 - 250470", "2 rows changed", null)],
                (entry, name, contact) =>
                {
                    prepared.Add((entry.BoardName, name, contact));

                    return Task.FromResult(entry.BoardName switch
                    {
                        "250425" => OtherDraftReady.Ready(Target("Commodore 64 - 250425", credited: true)),
                        "250470" => OtherDraftReady.Ready(Target("Commodore 64 - 250470")),
                        _ => OtherDraftReady.NotSent(SubmitDraftsWording.ErrorsFirst(2)),
                    });
                });

            window.SendOverrideForTests = (target, _) =>
            {
                sent.Add(target.DisplayName);

                return target.DisplayName.EndsWith("250470", StringComparison.Ordinal)
                    ? throw new HttpRequestException("no route")
                    : Task.FromResult(Accepted(sent.Count));
            };

            window.TickOtherDraftForTests(0);
            window.TickOtherDraftForTests(2);
            window.TickOtherDraftForTests(3);
            window.CreditNameTextBox.Text = "Bo";
            window.CreditContactTextBox.Text = "RHD";

            await window.SubmitNowAsync();

            Assert.Equal(
                [
                    "2 of 4 contributions sent",
                    "Commodore 64 - 250407: queued for review",
                    "Commodore 64 - 250425: queued for review",
                    "Commodore 64 - 250469: not sent - 2 errors to fix first",
                    "Commodore 64 - 250470: not sent - " + SubmitDraftsWording.NoConnection,
                ],
                window.OutcomeLinesForTests);

            Assert.Equal(1, window.Result!.SubmissionId);
        });

        Assert.Equal(["Commodore 64 - 250407", "Commodore 64 - 250425", "Commodore 64 - 250470"], sent);
        Assert.Equal([("250425", "Bo", "RHD"), ("250469", "Bo", "RHD"), ("250470", "Bo", "RHD")], prepared);

        // A draft credited under the name: the name is remembered, as this board's own credit is.
        Assert.Equal(("Bo", "RHD"), (UserSettings.CreditName, UserSettings.CreditContact));
    }

    // ###########################################################################################
    // The other drafts are credited only while that is wanted (code review, 2026-10-26): unticking
    // every credit of this board still put public credit rows naming the person into every other
    // draft sent. Unticking them all now unticks the other drafts' too - the name goes to none.
    // ###########################################################################################
    [Fact]
    public async Task Unticking_this_boards_every_credit_credits_none_of_the_other_drafts()
    {
        var prepared = new List<string>();

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(ownCredits: [new CreditSuggestion("Board data", "")]);
            window.OfferOtherDrafts(Offers, (_, name, _) =>
            {
                prepared.Add(name);
                return Task.FromResult(OtherDraftReady.Ready(Target("Commodore 64 - 250425")));
            });
            window.SendOverrideForTests = (_, _) => Task.FromResult(Accepted(1));

            window.TickOtherDraftForTests(0);
            window.CreditNameTextBox.Text = "Bo";
            Assert.All(window.CreditChoicesPanel.Children.OfType<CheckBox>(), box => box.IsChecked = false);
            Assert.False(window.CreditOthersCheckBox.IsChecked);

            await window.SubmitNowAsync();
        });

        Assert.Equal([string.Empty], prepared);
    }

    // ###########################################################################################
    // Credits put into the drafts before sending are said when nothing could be sent - the drafts are
    // not as they were (code review, 2026-10-26).
    // ###########################################################################################
    [Fact]
    public async Task Credits_put_into_the_other_drafts_are_said_when_nothing_could_be_sent()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.OfferOtherDrafts(Offers, (_, _, _) => Task.FromResult(OtherDraftReady.Ready(Target("Commodore 64 - 250425", credited: true))));
            window.SendOverrideForTests = (_, _) => throw new HttpRequestException("offline");

            window.TickOtherDraftForTests(0);
            window.CreditNameTextBox.Text = "Bo";
            await window.SubmitNowAsync();

            Assert.Equal("Contributions not sent", window.HeaderText.Text);
            Assert.Equal(SubmitDraftsWording.Intro(0, credited: true), window.GetControl<TextBlock>("OutcomeText").Text);
            Assert.Contains("credits", SubmitDraftsWording.Intro(0, credited: true), StringComparison.Ordinal);
        });
    }

    // A draft the server turned down is said so, its findings under it.
    [Fact]
    public async Task A_draft_the_server_turned_down_lists_what_it_found()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.OfferOtherDrafts(Offers, (_, _, _) => Task.FromResult(OtherDraftReady.Ready(Target("Commodore 64 - 250425"))));

            window.SendOverrideForTests = (target, _) => Task.FromResult(target.DisplayName.EndsWith("250425", StringComparison.Ordinal)
                ? new SubmissionResult
                {
                    SubmissionId = 2,
                    IsAccepted = false,
                    Findings = [new ValidationFinding { Severity = ValidationSeverity.Error, Subject = "U8", Message = "No category." }],
                }
                : Accepted(1));

            window.TickOtherDraftForTests(0);
            await window.SubmitNowAsync();

            Assert.Equal(
                [
                    "1 of 2 contributions sent",
                    "Commodore 64 - 250407: queued for review",
                    "Commodore 64 - 250425: not accepted - the server found problems to fix first",
                    "U8: No category.",
                ],
                window.OutcomeLinesForTests);
        });
    }

    // Nothing ticked: one board, said as ever - no list of boards, no credit name remembered.
    [Fact]
    public async Task With_nothing_ticked_one_board_is_sent_and_said_as_ever()
    {
        var sent = 0;

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.OfferOtherDrafts(Offers, (_, _, _) => throw new InvalidOperationException("not asked"));
            window.SendOverrideForTests = (_, _) =>
            {
                sent++;
                return Task.FromResult(Accepted(7));
            };

            await window.SubmitNowAsync();

            Assert.Equal("Contribution sent", window.HeaderText.Text);
            Assert.Equal(7, window.Result!.SubmissionId);
        });

        Assert.Equal(1, sent);
        Assert.Equal(string.Empty, UserSettings.CreditName);
    }
}
