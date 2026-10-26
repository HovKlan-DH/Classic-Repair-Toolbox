using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog's credit for the work sent (CRT 4.0 step 2, 2026-10-26 - the plan: "After a
// submission, CRT suggests a credit for what you sent"; owner decision: "Do it now"). Offered before
// anything is sent: each category the work deserves ticked, the name CRT remembers - and the ticked
// ones go into the draft first, so the board sent carries them. A blank name credits nobody.
//
// COLLECTION NOTE: "HeadlessUi" because it builds a Window; it points UserSettings at a temp file
// first (the dialog reads and remembers the credit name there), as MainWindowTests does.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowCreditsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SubmitDraftWindowCreditsTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static readonly CreditSuggestion[] Suggestions =
    [
        new("Schematic image", "Bottom"),
        new("Board data", ""),
    ];

    private SubmitDraftWindow BuildWindow(
        BoardData board,
        Func<IReadOnlyList<CreditEntry>, Task<(BoardData Board, string Fingerprint)?>>? addCredits,
        IReadOnlyList<CreditSuggestion>? suggestions = null)
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Commodore 64 - 250407",
            board,
            new SubmissionIdentity { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" },
            this.thisWorkspace.Path_("draft"),
            this.thisWorkspace.Path_("Data"),
            draftFingerprint: "before",
            creditSuggestions: suggestions ?? SubmitDraftWindowCreditsTests.Suggestions,
            addCredits: addCredits);

        return window;
    }

    [Fact]
    public void Each_credit_the_work_deserves_is_offered_ticked_with_the_name_remembered()
    {
        UserSettings.RememberCreditIdentity("Anna", "RHD");

        UiTest.Run(() =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), _ => Task.FromResult<(BoardData, string)?>(null));

            Assert.Equal(["Schematic image - Bottom", "Board data"], window.CreditChoicesForTests);
            Assert.All(window.GetControl<StackPanel>("CreditChoicesPanel").Children.OfType<CheckBox>(), box => Assert.True(box.IsChecked));
            Assert.Equal("Anna", window.GetControl<TextBox>("CreditNameTextBox").Text);
            Assert.Equal("RHD", window.GetControl<TextBox>("CreditContactTextBox").Text);
        });
    }

    // Nothing to credit, or no way to put a credit into the draft: no offer.
    [Fact]
    public void Without_suggestions_or_a_way_into_the_draft_nothing_is_offered()
    {
        UiTest.Run(() =>
        {
            Assert.False(this.BuildWindow(new BoardData(), null).GetControl<StackPanel>("CreditOfferPanel").IsVisible);
            Assert.False(this.BuildWindow(new BoardData(), _ => Task.FromResult<(BoardData, string)?>(null), []).GetControl<StackPanel>("CreditOfferPanel").IsVisible);
        });
    }

    // ###########################################################################################
    // The ticked credits go into the draft first, naming the person; the board and fingerprint then
    // sent are the draft's as it is afterwards - and the name and contact are remembered.
    // ###########################################################################################
    [Fact]
    public async Task The_ticked_credits_go_into_the_draft_first_and_the_name_is_remembered()
    {
        var withCredits = new BoardData { Credits = [new CreditEntry { Category = "Schematic image", NameOrHandle = "Bo" }] };
        var asked = new List<IReadOnlyList<CreditEntry>>();

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), credits =>
            {
                asked.Add(credits);
                return Task.FromResult<(BoardData, string)?>((withCredits, "after"));
            });

            window.GetControl<StackPanel>("CreditChoicesPanel").Children.OfType<CheckBox>().Last().IsChecked = false;
            window.GetControl<TextBox>("CreditNameTextBox").Text = " Bo ";
            window.GetControl<TextBox>("CreditContactTextBox").Text = "https://example.org";

            Assert.True(await window.AddCreditsFirstAsync());

            CreditEntry credit = Assert.Single(Assert.Single(asked));
            Assert.Equal(("Schematic image", "Bottom", "Bo", "https://example.org"), (credit.Category, credit.SubCategory, credit.NameOrHandle, credit.Contact));
            Assert.Equal("after", window.NewReceipt(1, "token", DateTimeOffset.UtcNow).DraftFingerprint);

            // In: a second try adds nothing again.
            Assert.True(await window.AddCreditsFirstAsync());
            Assert.Single(asked);
        });

        Assert.Equal(("Bo", "https://example.org"), (UserSettings.CreditName, UserSettings.CreditContact));
    }

    // ###########################################################################################
    // A credit is put into the draft BEFORE sending - so a send that then fails says so (code review,
    // 2026-10-26): it said "nothing has been lost, and your draft is unchanged" over a draft that now
    // held the credit rows.
    // ###########################################################################################
    [Fact]
    public async Task A_credit_put_into_the_draft_is_said_when_the_send_then_fails()
    {
        var withCredits = new BoardData { Credits = [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }] };

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), _ => Task.FromResult<(BoardData, string)?>((withCredits, "after")));
            window.GetControl<TextBox>("CreditNameTextBox").Text = "Bo";
            window.SendOverrideForTests = (_, _) => throw new System.Net.Http.HttpRequestException("offline");

            await window.SubmitNowAsync();

            string said = window.GetControl<TextBlock>("OutcomeText").Text ?? string.Empty;
            Assert.DoesNotContain("your draft is unchanged", said, StringComparison.Ordinal);
            Assert.Contains("Your credit is in your draft now", said, StringComparison.Ordinal);
        });
    }

    // Without a credit added, the draft truly is unchanged - and says so as before.
    [Fact]
    public async Task Without_a_credit_a_send_that_fails_says_the_draft_is_unchanged()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), _ => Task.FromResult<(BoardData, string)?>(null));
            window.GetControl<TextBox>("CreditNameTextBox").Text = " ";
            window.SendOverrideForTests = (_, _) => throw new System.Net.Http.HttpRequestException("offline");

            await window.SubmitNowAsync();

            Assert.Contains("your draft is unchanged", window.GetControl<TextBlock>("OutcomeText").Text ?? string.Empty, StringComparison.Ordinal);
        });
    }

    // A blank name credits nobody - nothing goes into the draft, and the submission goes ahead.
    [Fact]
    public async Task A_blank_name_credits_nobody()
    {
        var asked = 0;

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), _ =>
            {
                asked++;
                return Task.FromResult<(BoardData, string)?>(null);
            });

            window.GetControl<TextBox>("CreditNameTextBox").Text = "  ";

            Assert.True(await window.AddCreditsFirstAsync());
        });

        Assert.Equal(0, asked);
        Assert.Equal(string.Empty, UserSettings.CreditName);
    }

    // The draft could not take them: false - the dialog sends nothing - and no name is remembered.
    [Fact]
    public async Task Credits_the_draft_could_not_take_stop_the_submission()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData(), _ => Task.FromResult<(BoardData, string)?>(null));
            window.GetControl<TextBox>("CreditNameTextBox").Text = "Bo";

            Assert.False(await window.AddCreditsFirstAsync());
        });

        Assert.Equal(string.Empty, UserSettings.CreditName);
        Assert.Contains("nothing was sent", SubmitDraftWindow.CreditsNotAdded, StringComparison.Ordinal);
    }
}
