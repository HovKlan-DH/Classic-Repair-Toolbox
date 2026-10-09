using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > "Order of boards", BoardOrderView (owner request, 2026-10-04: "a possibility to be able
// to sort the list of systems, which then gets saved to both sources (BETA + stable) after my save").
// Drawn without a server: the list through UseListing, a move as a drag would make it, and Save
// answered by the test - which also sees the order sent.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardOrderViewTests
{
    private static BoardListingAnswer Listing() =>
        new(
            true,
            [
                new BoardListingRow("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                new BoardListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                new BoardListingRow("ZX Spectrum/Spectrum 16K-48K/Issue 4B", "ZX Spectrum 16K/48K", "Issue 4B", "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Issue 4B v2.0.0.xlsx"),
            ],
            []);

    private static BoardOrderView Shown()
    {
        var view = new BoardOrderView();
        view.UseListing(BoardOrderViewTests.Listing());
        return view;
    }

    // The whole list, in CRT's order - and nothing to save until something moved.
    [Fact]
    public void The_list_shows_in_CRTs_order_with_nothing_to_save()
    {
        UiTest.Run(() =>
        {
            BoardOrderView view = BoardOrderViewTests.Shown();

            Assert.Equal(["250407 (long board)", "310378", "Issue 4B"], view.Rows.Select(row => row.BoardName));
            Assert.False(view.IsReordered);
            Assert.False(view.IsSaveEnabledForTests);
            Assert.Equal(string.Empty, view.MessageForTests);
        });
    }

    // ###########################################################################################
    // *** "HARDWARE" AND "BOARD" OVER THE TWO COLUMNS (owner request, 2026-10-09: "I would like to
    // have a header for 'Hardware' and 'Board' for the table"), each starting where its column's
    // names start in every row.
    // ###########################################################################################
    [Fact]
    public void The_list_has_Hardware_and_Board_headings_over_its_two_columns()
    {
        UiTest.Run(() =>
        {
            BoardOrderView view = BoardOrderViewTests.Shown();
            var window = new Window { Content = view, Width = 900, Height = 500 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            TextBlock hardware = view.GetControl<TextBlock>("HardwareHeadingText");
            TextBlock board = view.GetControl<TextBlock>("BoardHeadingText");
            Assert.Equal(("Hardware", "Board"), (hardware.Text, board.Text));

            List<TextBlock> first = view.GetVisualDescendants()
                .OfType<TextBlock>()
                .Where(text => text.Text is "Commodore 64" or "250407 (long board)")
                .ToList();
            double X(Visual visual) => visual.TranslatePoint(default, view)!.Value.X;

            Assert.Equal(X(first.Single(text => text.Text == "Commodore 64")), X(hardware), precision: 0);
            Assert.Equal(X(first.Single(text => text.Text == "250407 (long board)")), X(board), precision: 0);

            window.Close();
        });
    }

    // ###########################################################################################
    // A move turns Save on and says the order is not saved yet; "Cancel" drops it again (the button
    // read "Put back as saved" until the owner asked for "Cancel", 2026-10-04).
    // ###########################################################################################
    [Fact]
    public void A_move_offers_Save_and_putting_it_back_drops_it()
    {
        UiTest.Run(() =>
        {
            BoardOrderView view = BoardOrderViewTests.Shown();

            Assert.Equal("Cancel", view.FindControl<Button>("UndoButton")!.Content);

            view.MoveForTests(2, 0);

            Assert.True(view.IsReordered);
            Assert.True(view.IsSaveEnabledForTests);
            Assert.Equal(BoardOrderDisplay.Unsaved, view.MessageForTests);

            view.PutBack();

            Assert.Equal(["250407 (long board)", "310378", "Issue 4B"], view.Rows.Select(row => row.BoardName));
            Assert.False(view.IsSaveEnabledForTests);
            Assert.Equal(string.Empty, view.MessageForTests);
        });
    }

    // ###########################################################################################
    // Save sends EVERY board, in the order on screen, and says what the server did. What was sent
    // is then the list as saved - nothing waits to be saved any more.
    // ###########################################################################################
    [Fact]
    public async Task Saving_sends_every_board_in_the_new_order_and_says_what_was_done()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardOrderView view = BoardOrderViewTests.Shown();
            IReadOnlyList<string>? sent = null;

            view.SaveOverrideForTests = order =>
            {
                sent = order;
                return Task.FromResult(ReviewApiResult<BoardOrderAnswer>.Ok(new BoardOrderAnswer(true, true)));
            };

            view.MoveForTests(2, 0);
            await view.SaveAsync();

            Assert.Equal(["ZX Spectrum/Spectrum 16K-48K/Issue 4B", "Commodore/C64/250407", "Commodore/C128/310378"], sent);
            Assert.Equal(BoardOrderDisplay.Saved(new BoardOrderAnswer(true, true)).Text, view.MessageForTests);
            Assert.False(view.IsReordered);
            Assert.False(view.IsSaveEnabledForTests);
        });
    }

    // A board on two rows of the list goes to the server once - the server refuses it twice
    // (code review, 2026-10-04).
    [Fact]
    public async Task A_board_listed_on_two_rows_is_sent_once()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardOrderView();
            view.UseListing(new BoardListingAnswer(
                true,
                [
                    new BoardListingRow("Commodore/C64/250407", "Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                    new BoardListingRow("Commodore/C64/250407", "Commodore 64", "250407 (PAL)", "Commodore/C64/250407/Data C64 250407 PAL v2.0.0.xlsx"),
                    new BoardListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                ],
                []));

            IReadOnlyList<string>? sent = null;

            view.SaveOverrideForTests = order =>
            {
                sent = order;
                return Task.FromResult(ReviewApiResult<BoardOrderAnswer>.Ok(new BoardOrderAnswer(true, true)));
            };

            view.MoveForTests(2, 0);
            await view.SaveAsync();

            Assert.Equal(["Commodore/C128/310378", "Commodore/C64/250407"], sent);
        });
    }

    // A refusal - the list changed meanwhile - keeps the moves on screen with the server's words.
    [Fact]
    public async Task A_refused_save_keeps_the_moves_and_shows_the_servers_words()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardOrderView view = BoardOrderViewTests.Shown();

            view.SaveOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<BoardOrderAnswer>.Failed(ReviewApiFailure.Conflict, "The list in BETA has changed since you opened it."));

            view.MoveForTests(2, 0);
            await view.SaveAsync();

            Assert.Equal("The list in BETA has changed since you opened it.", view.MessageForTests);
            Assert.True(view.IsReordered);
            Assert.Equal("Issue 4B", view.Rows[0].BoardName);
        });
    }

    // No versioned main Excel data file in BETA: said, never an empty box that reads as "no boards".
    [Fact]
    public void Without_a_list_it_says_there_is_none()
    {
        UiTest.Run(() =>
        {
            var view = new BoardOrderView();
            view.UseListing(new BoardListingAnswer(false, [], []));

            Assert.Empty(view.Rows);
            Assert.Equal(BoardOrderDisplay.NoList, view.MessageForTests);
        });
    }

    // Signed out: nothing of the list stays.
    [Fact]
    public void Clearing_leaves_nothing()
    {
        UiTest.Run(() =>
        {
            BoardOrderView view = BoardOrderViewTests.Shown();
            view.MoveForTests(2, 0);

            view.Clear();

            Assert.Empty(view.Rows);
            Assert.False(view.IsReordered);
            Assert.False(view.IsSaveEnabledForTests);
        });
    }
}
