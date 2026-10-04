using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > "Order of systems", SystemOrderView (owner request, 2026-10-04: "a possibility to be able
// to sort the list of systems, which then gets saved to both sources (BETA + stable) after my save").
// Drawn without a server: the list through UseListing, a move as a drag would make it, and Save
// answered by the test - which also sees the order sent.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemOrderViewTests
{
    private static SystemListingAnswer Listing() =>
        new(
            true,
            [
                new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                new SystemListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                new SystemListingRow("ZX Spectrum/Spectrum 16K-48K/Issue 4B", "ZX Spectrum 16K/48K", "Issue 4B", "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Issue 4B v2.0.0.xlsx"),
            ],
            []);

    private static SystemOrderView Shown()
    {
        var view = new SystemOrderView();
        view.UseListing(SystemOrderViewTests.Listing());
        return view;
    }

    // The whole list, in CRT's order - and nothing to save until something moved.
    [Fact]
    public void The_list_shows_in_CRTs_order_with_nothing_to_save()
    {
        UiTest.Run(() =>
        {
            SystemOrderView view = SystemOrderViewTests.Shown();

            Assert.Equal(["250407 (long board)", "310378", "Issue 4B"], view.Rows.Select(row => row.BoardName));
            Assert.False(view.IsReordered);
            Assert.False(view.IsSaveEnabledForTests);
            Assert.Equal(string.Empty, view.MessageForTests);
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
            SystemOrderView view = SystemOrderViewTests.Shown();

            Assert.Equal("Cancel", view.FindControl<Button>("UndoButton")!.Content);

            view.MoveForTests(2, 0);

            Assert.True(view.IsReordered);
            Assert.True(view.IsSaveEnabledForTests);
            Assert.Equal(SystemOrderDisplay.Unsaved, view.MessageForTests);

            view.PutBack();

            Assert.Equal(["250407 (long board)", "310378", "Issue 4B"], view.Rows.Select(row => row.BoardName));
            Assert.False(view.IsSaveEnabledForTests);
            Assert.Equal(string.Empty, view.MessageForTests);
        });
    }

    // ###########################################################################################
    // Save sends EVERY system, in the order on screen, and says what the server did. What was sent
    // is then the list as saved - nothing waits to be saved any more.
    // ###########################################################################################
    [Fact]
    public async Task Saving_sends_every_system_in_the_new_order_and_says_what_was_done()
    {
        await UiTest.RunAsync(async () =>
        {
            SystemOrderView view = SystemOrderViewTests.Shown();
            IReadOnlyList<string>? sent = null;

            view.SaveOverrideForTests = order =>
            {
                sent = order;
                return Task.FromResult(ReviewApiResult<SystemOrderAnswer>.Ok(new SystemOrderAnswer(true, true)));
            };

            view.MoveForTests(2, 0);
            await view.SaveAsync();

            Assert.Equal(["ZX Spectrum/Spectrum 16K-48K/Issue 4B", "Commodore/C64/250407", "Commodore/C128/310378"], sent);
            Assert.Equal(SystemOrderDisplay.Saved(new SystemOrderAnswer(true, true)).Text, view.MessageForTests);
            Assert.False(view.IsReordered);
            Assert.False(view.IsSaveEnabledForTests);
        });
    }

    // A system on two rows of the list goes to the server once - the server refuses it twice
    // (code review, 2026-10-04).
    [Fact]
    public async Task A_system_listed_on_two_rows_is_sent_once()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new SystemOrderView();
            view.UseListing(new SystemListingAnswer(
                true,
                [
                    new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                    new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407 (PAL)", "Commodore/C64/250407/Data C64 250407 PAL v2.0.0.xlsx"),
                    new SystemListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                ],
                []));

            IReadOnlyList<string>? sent = null;

            view.SaveOverrideForTests = order =>
            {
                sent = order;
                return Task.FromResult(ReviewApiResult<SystemOrderAnswer>.Ok(new SystemOrderAnswer(true, true)));
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
            SystemOrderView view = SystemOrderViewTests.Shown();

            view.SaveOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<SystemOrderAnswer>.Failed(ReviewApiFailure.Conflict, "The list in BETA has changed since you opened it."));

            view.MoveForTests(2, 0);
            await view.SaveAsync();

            Assert.Equal("The list in BETA has changed since you opened it.", view.MessageForTests);
            Assert.True(view.IsReordered);
            Assert.Equal("Issue 4B", view.Rows[0].BoardName);
        });
    }

    // No versioned main Excel data file in BETA: said, never an empty box that reads as "no systems".
    [Fact]
    public void Without_a_list_it_says_there_is_none()
    {
        UiTest.Run(() =>
        {
            var view = new SystemOrderView();
            view.UseListing(new SystemListingAnswer(false, [], []));

            Assert.Empty(view.Rows);
            Assert.Equal(SystemOrderDisplay.NoList, view.MessageForTests);
        });
    }

    // Signed out: nothing of the list stays.
    [Fact]
    public void Clearing_leaves_nothing()
    {
        UiTest.Run(() =>
        {
            SystemOrderView view = SystemOrderViewTests.Shown();
            view.MoveForTests(2, 0);

            view.Clear();

            Assert.Empty(view.Rows);
            Assert.False(view.IsReordered);
            Assert.False(view.IsSaveEnabledForTests);
        });
    }
}
