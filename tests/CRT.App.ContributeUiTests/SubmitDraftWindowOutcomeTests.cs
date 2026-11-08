using System.Net.Http;
using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// What the Submit dialog says when a send ends, and while it runs (code review, 2026-11-02).
//
//   - A send that never got an answer says why in its own sentence and adds nothing: every such failure
//     used to end "The server did not say what was wrong, which is a fault in CRT rather than in your
//     data" - beside "Check your internet connection". That apology is for what the server REFUSED.
//   - A file that is there but cannot be used is said for what it is, not "moved or deleted".
//
// Nothing goes over the network (test rule 6): SendOverrideForTests stands in for a send, and a file
// that cannot be used stops the real one before its first request.
//
// COLLECTION NOTE: "HeadlessUi" because it builds a Window; UserSettings points at a temp file (the
// dialog reads the address and the credit name there).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowOutcomeTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SubmitDraftWindowOutcomeTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private SubmitDraftWindow BuildWindow(
        BoardData? board = null,
        IReadOnlyList<CreditSuggestion>? suggestions = null,
        Func<IReadOnlyList<CreditEntry>, Task<(BoardData Board, string Fingerprint)?>>? addCredits = null)
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Commodore 64 - 250407",
            board ?? new BoardData(),
            new SubmissionIdentity { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" },
            Directory.CreateDirectory(Path.Combine(this.thisWorkspace.Root, "draft")).FullName,
            Directory.CreateDirectory(Path.Combine(this.thisWorkspace.Root, "Data")).FullName,
            draftFingerprint: "before",
            creditSuggestions: suggestions,
            addCredits: addCredits);

        return window;
    }

    private static string Said(SubmitDraftWindow window) => window.GetControl<TextBlock>("OutcomeText").Text ?? string.Empty;

    private static List<string> Findings(SubmitDraftWindow window) =>
        window.GetControl<StackPanel>("FindingsPanel").Children.OfType<TextBlock>().Select(text => text.Text ?? string.Empty).ToList();

    // ------------------------------------------------------------------ the apology is for a refusal

    [Fact]
    public async Task A_send_that_never_got_an_answer_says_only_that()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.SendOverrideForTests = (_, _) => throw new HttpRequestException("offline");

            await window.SubmitNowAsync();

            Assert.Equal(SubmitDraftWindow.NoConnectionMessage(credited: false), Said(window));
            Assert.Empty(Findings(window));
        });
    }

    [Fact]
    public async Task Something_unexpected_says_only_that()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.SendOverrideForTests = (_, _) => throw new InvalidOperationException("no idea");

            await window.SubmitNowAsync();

            Assert.Equal(SubmitDraftWindow.UnexpectedMessage(credited: false), Said(window));
            Assert.Empty(Findings(window));
        });
    }

    // An upload refused, or a sign-in page's answer: the sentence the exception carries says what happened.
    [Theory]
    [InlineData("Uploading [top.png] failed: The server is short of storage space.")]
    [InlineData(SubmissionClient.UnreadableReply)]
    public async Task A_refusal_in_a_sentence_says_only_that(string sentence)
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.SendOverrideForTests = (_, _) => throw new SubmissionRejectedException(0, [], sentence);

            await window.SubmitNowAsync();

            Assert.Equal(sentence, Said(window));
            Assert.Empty(Findings(window));
        });
    }

    [Fact]
    public async Task A_credit_the_draft_could_not_take_says_only_that()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(
                suggestions: [new CreditSuggestion("Board data", "")],
                addCredits: _ => Task.FromResult<(BoardData, string)?>(null));

            window.GetControl<TextBox>("CreditNameTextBox").Text = "Bo";

            await window.SubmitNowAsync();

            Assert.Equal(SubmitDraftWindow.CreditsNotAdded, Said(window));
            Assert.Empty(Findings(window));
        });
    }

    // What came with a failed send is still shown.
    [Fact]
    public async Task A_refusal_that_carries_findings_lists_them()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.SendOverrideForTests = (_, _) => throw new SubmissionRejectedException(
                400,
                [new ValidationFinding { Severity = ValidationSeverity.Error, Subject = "U8", Message = "No category." }],
                "The server refused the submission (400).");

            await window.SubmitNowAsync();

            Assert.Equal(["U8: No category."], Findings(window));
        });
    }

    // The apology stays where it is true: the server answered, refused, and gave no reason.
    [Fact]
    public async Task A_result_the_server_refused_with_no_reason_still_says_so()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow();
            window.SendOverrideForTests = (_, _) => Task.FromResult(new SubmissionResult { SubmissionId = 1, IsAccepted = false, State = "rejected" });

            await window.SubmitNowAsync();

            Assert.Contains("fault in CRT", Assert.Single(Findings(window)), StringComparison.Ordinal);
        });
    }

    // ------------------------------------------------------------------ the file that cannot be used

    // ###########################################################################################
    // *** THE REASON A FILE CANNOT BE USED IS SAID (code review, 2026-11-02). *** Hashing the files
    // collected "[x] cannot be used: ..." and "[x] could not be read: ...", and nothing read them: the
    // contributor was told the file "could not be found ... moved or deleted" for a file that was there.
    // Through the real send, which stops on these before its first request.
    // ###########################################################################################
    [Fact]
    public async Task A_name_CRT_refuses_is_said_for_what_it_is_not_as_a_file_moved_or_deleted()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "../secret.png" }],
            });

            await window.SubmitNowAsync();

            string finding = Assert.Single(Findings(window));
            Assert.StartsWith("[../secret.png] cannot be used: ", finding, StringComparison.Ordinal);
            Assert.DoesNotContain("could not be found", finding, StringComparison.Ordinal);
            Assert.DoesNotContain("missing", Said(window), StringComparison.Ordinal);
        });
    }

    [Fact]
    public async Task A_file_that_is_there_but_cannot_be_read_says_so()
    {
        string file = this.thisWorkspace.WriteFile(Path.Combine("draft", "top.png"), "picture");

        await UiTest.RunAsync(async () =>
        {
            SubmitDraftWindow window = this.BuildWindow(new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" }],
            });

            using (new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                await window.SubmitNowAsync();
            }

            string finding = Assert.Single(Findings(window));
            Assert.StartsWith("[top.png] could not be read: ", finding, StringComparison.Ordinal);
            Assert.DoesNotContain("moved or deleted", finding, StringComparison.Ordinal);
        });
    }

    // ------------------------------------------------------------------ while it runs

    // The dialog is modal over CRT (TabDrafts.SubmitAsync), so "You can keep using CRT" was untrue.
    [Fact]
    public void The_progress_panel_promises_nothing_the_modal_dialog_does_not_do()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();

            List<string> lines = window.GetControl<StackPanel>("ProgressPanel").Children
                .OfType<TextBlock>()
                .Select(text => text.Text ?? string.Empty)
                .ToList();

            Assert.DoesNotContain(lines, line => line.Contains("keep using CRT", StringComparison.Ordinal));
            Assert.Contains(lines, line => line.StartsWith("Cancelling is safe", StringComparison.Ordinal));
        });
    }
}
