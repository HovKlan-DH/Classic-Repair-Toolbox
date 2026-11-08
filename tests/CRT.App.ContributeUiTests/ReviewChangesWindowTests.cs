using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Review changes" (CRT 4.0 step 1's last slice, 2026-10-25 - the plan: "Review changes shows a
// readable summary ... Undo changes to IC1 - one component, back to its published values"). The
// window shows DraftChangeSummary's groups as they are, offers each COMPONENT its own red "Undo
// changes to [..]" (never the board's group), and after an undo shows the summary it is handed back
// - or why the component could not be put back.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ReviewChangesWindowTests
{
    private static BoardContent Content(params ComponentEntry[] components) =>
        new(new BoardData { Components = [.. components] }, [], []);

    // U8's description changed and U9 added - two component groups.
    private static DraftChangeSummary TwoComponents() =>
        DraftChangeSummary.Build(
            ReviewChangesWindowTests.Content(new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Description = "old" }),
            ReviewChangesWindowTests.Content(
                new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Description = "new" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "VIC" }));

    // Only U9 left - what the summary reads after U8 was put back.
    private static DraftChangeSummary OnlyU9() =>
        DraftChangeSummary.Build(
            ReviewChangesWindowTests.Content(new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" }),
            ReviewChangesWindowTests.Content(
                new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "VIC" }));

    [Fact]
    public void Each_component_is_a_group_with_its_own_red_undo_button()
    {
        UiTest.Run(() =>
        {
            var window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", ReviewChangesWindowTests.TwoComponents(), _ => Task.FromResult((DraftChangeSummary.Empty, (string?)null)));

            Assert.Equal(["U8 - CPU", "U9 - VIC"], window.GroupTitlesForTests);
            Assert.False(window.ShowsNothingForTests);

            IReadOnlyList<Button> buttons = window.UndoButtonsForTests;
            Assert.Equal(["Undo changes to [U8]", "Undo changes to [U9]"], buttons.Select(button => button.Content as string));

            // It throws the component's changes away - red (CLAUDE.md, button colours), and no "...".
            Assert.All(buttons, button => Assert.Contains("Danger", button.Classes));
            Assert.All(buttons, button => Assert.DoesNotContain("...", (string)button.Content!, StringComparison.Ordinal));
        });
    }

    // The board's own group (renames, calibrations, KiCad data) has no "Undo changes to".
    [Fact]
    public void The_boards_own_group_offers_no_undo()
    {
        UiTest.Run(() =>
        {
            var published = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }] };
            var draft = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "Images/top.png" }] };

            var window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", DraftChangeSummary.Build(new BoardContent(published, [], []), new BoardContent(draft, [], [])), _ => Task.FromResult((DraftChangeSummary.Empty, (string?)null)));

            Assert.Equal([DraftChangeSummary.BoardTitle], window.GroupTitlesForTests);
            Assert.Empty(window.UndoButtonsForTests);
        });
    }

    // With nowhere to send an undo, none is offered.
    [Fact]
    public void Without_an_undo_to_call_no_button_is_offered()
    {
        UiTest.Run(() =>
        {
            var window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", ReviewChangesWindowTests.TwoComponents());

            Assert.Equal(2, window.GroupTitlesForTests.Count);
            Assert.Empty(window.UndoButtonsForTests);
        });
    }

    [Fact]
    public void A_draft_changing_nothing_says_so()
    {
        UiTest.Run(() =>
        {
            var window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", DraftChangeSummary.Empty);

            Assert.True(window.ShowsNothingForTests);
            Assert.Empty(window.GroupTitlesForTests);
        });
    }

    private static string IntroOf(ReviewChangesWindow window) => window.GetControl<TextBlock>("IntroText").Text ?? string.Empty;

    private static string NothingOf(ReviewChangesWindow window) => window.GetControl<TextBlock>("NothingText").Text ?? string.Empty;

    // ###########################################################################################
    // *** A NEW BOARD HAS NO CHANGES TO LIST (owner request, 2026-11-03: "the only thing it needs to state
    // is 'New board - not submitted yet' ... viewing 'a million' changes on a new board is not helpful").
    // ***
    // ###########################################################################################
    [Theory]
    [InlineData(null, "New board - not submitted yet")]
    [InlineData("submitted", "New board - submitted")]
    [InlineData("published", "New board - published")]
    public void A_new_boards_window_says_where_it_stands_and_lists_nothing(string? sentState, string words)
    {
        UiTest.Run(() =>
        {
            var window = new ReviewChangesWindow();
            SubmissionReceipt[] sent = sentState is null ? [] : [new SubmissionReceipt { SubmissionId = 1, LastKnownState = sentState }];

            window.InitializeNewBoard("Commodore 64 - 999001", ReviewChangesWindow.NewBoardStanding(sent));

            Assert.StartsWith(words, IntroOf(window), StringComparison.Ordinal);
            Assert.Empty(window.GroupTitlesForTests);
            Assert.Empty(window.UndoButtonsForTests);
            Assert.False(window.ShowsNothingForTests);
        });
    }

    [Fact]
    public void A_published_boards_window_says_what_it_compares_with_and_explains_the_undo()
    {
        UiTest.Run(() =>
        {
            var window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", ReviewChangesWindowTests.TwoComponents(), _ => Task.FromResult((DraftChangeSummary.Empty, (string?)null)));

            Assert.Contains("compared with the published data", IntroOf(window), StringComparison.Ordinal);
            Assert.Contains("\"Undo changes to\"", IntroOf(window), StringComparison.Ordinal);
            Assert.Contains("published data", NothingOf(window), StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // Undo changes to [U8]: the host is asked for THAT label, and the summary it answers is shown -
    // U8's group goes, U9's stays.
    // ###########################################################################################
    [Fact]
    public async Task Undoing_a_component_asks_for_its_label_and_shows_the_summary_answered()
    {
        var asked = new List<string>();

        ReviewChangesWindow window = null!;

        UiTest.Run(() =>
        {
            window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", ReviewChangesWindowTests.TwoComponents(), label =>
            {
                asked.Add(label);
                return Task.FromResult((ReviewChangesWindowTests.OnlyU9(), (string?)null));
            });
        });

        await UiTest.RunAsync(() => window.UndoComponentAsync("U8"));

        UiTest.Run(() =>
        {
            Assert.Equal(["U8"], asked);
            Assert.Equal(["U9 - VIC"], window.GroupTitlesForTests);
            Assert.Equal(string.Empty, window.StatusTextForTests);
        });
    }

    // One that could not be put back says why, over the summary as it now is.
    [Fact]
    public async Task A_component_that_could_not_be_put_back_says_why()
    {
        ReviewChangesWindow window = null!;

        UiTest.Run(() =>
        {
            window = new ReviewChangesWindow();
            window.Initialize("Commodore 64 - 250407", ReviewChangesWindowTests.TwoComponents(), _ =>
                Task.FromResult((ReviewChangesWindowTests.TwoComponents(), (string?)DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.WriteFailed))));
        });

        await UiTest.RunAsync(() => window.UndoComponentAsync("U8"));

        UiTest.Run(() =>
        {
            Assert.Equal("[U8] could not be put back. " + DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.WriteFailed), window.StatusTextForTests);
            Assert.Equal(["U8 - CPU", "U9 - VIC"], window.GroupTitlesForTests);
        });
    }
}
