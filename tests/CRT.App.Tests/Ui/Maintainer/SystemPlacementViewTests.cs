using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Placing a new system in CRT's drop-down lists, on the Systems screen (owner request, 2026-09-27:
// "The maintainer should order the new system, so it becomes visible in the right location for
// the drop-down lists. This must be done before it can be pushed to BETA.").
//
// The panel shows BETA's whole list with the new system as the one panel that moves - dragged with
// a REAL pointer here, through CRT.UI's ListRowDrag (the Drafts tab's schematic images' drag) - and
// Save sends the row above it. Save is answered without a server (SaveOverrideForTests). The rules
// and words are SystemPlacementDisplay's and tested there.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemPlacementViewTests
{
    private const string C64 = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";
    private const string C128 = "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx";
    private const string C128Dcr = "Commodore/C128/250477/Data C128DCR 250477 v2.0.0.xlsx";
    private const string Spectrum = "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Issue 4B v2.0.0.xlsx";

    private const string Open128 = "Commodore/C128/310378 Open128";

    private static readonly IReadOnlyList<SystemListingRow> Listed =
    [
        new("Commodore/C64/250407", "Commodore 64", "250407 (long board)", C64),
        new("Commodore/C128/310378", "Commodore 128", "310378 (C128 & C128D)", C128),
        new("Commodore/C128/250477", "Commodore 128", "250477 (C128DCR)", C128Dcr),
        new("ZX Spectrum/Spectrum 16K-48K/Issue 4B", "ZX Spectrum 16K/48K", "Issue 4B", Spectrum),
    ];

    // Never placed; the server suggests it after the last C128 board.
    private static UnlistedSystemEntry Entry(SystemPlacement? placement = null, bool canPlace = true) =>
        new(Open128, "Commodore", "C128", "310378 Open128", false, canPlace, placement,
            new SystemPlacement("Commodore 128", "310378 Open128", string.Empty, C128Dcr));

    private static SystemListingAnswer Listing(UnlistedSystemEntry? entry = null, IReadOnlyList<SystemListingRow>? listed = null) =>
        new(true, listed ?? SystemPlacementViewTests.Listed, [entry ?? SystemPlacementViewTests.Entry()]);

    // The lists as a later check finds them: another system has been listed meanwhile - a real
    // change, so only the unsaved-changes rule can keep the panel as it is.
    private static SystemListingAnswer ChangedListing() =>
        SystemPlacementViewTests.Listing(listed:
        [
            .. SystemPlacementViewTests.Listed,
            new SystemListingRow("Amiga/A500/Rev 6A", "Amiga 500", "Rev 6A", "Amiga/A500/Rev 6A/Data A500 Rev 6A v2.0.0.xlsx"),
        ]);

    // Shown in a real window, tall enough that every row is on screen - a press on a row scrolled
    // out of sight lands on nothing.
    private static (Window Window, SystemPlacementView View) Shown(SystemListingAnswer listing)
    {
        var view = new SystemPlacementView();
        var window = new Window { Content = view, Width = 900, Height = 1000 };

        window.Show();
        view.Show(listing, Open128);
        Dispatcher.UIThread.RunJobs();

        return (window, view);
    }

    private static List<string> Boards(SystemPlacementView view) => view.Rows.Select(row => row.BoardName).ToList();

    private static Control RowContainer(SystemPlacementView view, int index) =>
        view.GetControl<ItemsControl>("PlacementItemsControl").ContainerFromIndex(index)!;

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    private static void Press(Window window, Point point)
    {
        window.MouseDown(point, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static void Move(Window window, Point point)
    {
        window.MouseMove(point, Avalonia.Input.RawInputModifiers.LeftMouseButton);
        Dispatcher.UIThread.RunJobs();
    }

    private static void Release(Window window, Point point)
    {
        window.MouseUp(point, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    // Save, answered as the server would, keeping what was sent.
    private static List<SetPlacementRequest> AnswerSaves(SystemPlacementView view, bool listedInBeta = false)
    {
        var sent = new List<SetPlacementRequest>();

        view.SaveOverrideForTests = request =>
        {
            sent.Add(request);
            return Task.FromResult(ReviewApiResult<SetPlacementAnswer>.Ok(new SetPlacementAnswer(
                new SystemPlacement(request.HardwareName!, request.BoardName!, request.Notes ?? string.Empty, request.AfterExcelDataFile),
                listedInBeta,
                "Its place is saved.")));
        };

        return sent;
    }

    [Fact]
    public void The_new_system_is_shown_in_the_whole_list_where_it_was_suggested_with_its_names()
    {
        UiTest.Run(() =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());

            Assert.True(view.IsVisible);
            Assert.Equal(
                ["250407 (long board)", "310378 (C128 & C128D)", "250477 (C128DCR)", "310378 Open128", "Issue 4B"],
                SystemPlacementViewTests.Boards(view));

            // Only the system being placed is its own panel, and only it moves.
            Assert.Equal([false, false, false, true, false], view.Rows.Select(row => row.CanDrag));
            Assert.True(view.Rows[3].ShowsAsNewSystem);

            Assert.Equal("Commodore 128 is hardware 2 of 3 in CRT, and 310378 Open128 its board 3 of 3.", view.TextOfForTests("WhereInCrtText"));
            Assert.Equal("No place saved yet - the one shown is a suggestion.", view.TextOfForTests("SavedStateText"));

            // Accepting the suggestion needs no change: Save is ready at once.
            Assert.True(view.IsSaveEnabledForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A REAL POINTER DRAG, AND WHAT SAVE SENDS. *** Press the panel, move it up to the first row,
    // release: it is first, and Save sends "after nothing". The dashed placeholder stands in its
    // place while it moves - the same one the schematic images show.
    // ###########################################################################################
    [Fact]
    public async Task Dragging_the_panel_to_the_top_places_it_first_and_Save_sends_no_row_above_it()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());
            List<SetPlacementRequest> sent = SystemPlacementViewTests.AnswerSaves(view);

            Point panel = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 3));
            Point top = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 0)) - new Point(0, 6);

            SystemPlacementViewTests.Press(window, panel);
            SystemPlacementViewTests.Move(window, top);

            PlacementListRow moving = view.Rows[0];
            Assert.True(moving.IsNewSystem);
            Assert.True(moving.IsDropPlaceholder);

            SystemPlacementViewTests.Release(window, top);

            Assert.False(moving.IsDropPlaceholder);
            Assert.Equal("310378 Open128", SystemPlacementViewTests.Boards(view)[0]);
            Assert.Equal("Commodore 128 is hardware 1 of 3 in CRT, and 310378 Open128 its board 1 of 3.", view.TextOfForTests("WhereInCrtText"));

            await view.SaveAsync();

            SetPlacementRequest request = Assert.Single(sent);
            Assert.Equal(Open128, request.SystemId);
            Assert.Equal("Commodore 128", request.HardwareName);
            Assert.Equal("310378 Open128", request.BoardName);
            Assert.Null(request.AfterExcelDataFile);
            Assert.Equal("Its place is saved.", view.TextOfForTests("MessageText"));

            window.Close();
        });
    }

    [Fact]
    public async Task Dropped_between_two_rows_Save_sends_the_row_above_it()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());
            List<SetPlacementRequest> sent = SystemPlacementViewTests.AnswerSaves(view);

            Point panel = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 3));

            // Just above the C128's middle: below the C64, so straight after it.
            Point target = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 1)) - new Point(0, 3);

            SystemPlacementViewTests.Press(window, panel);
            SystemPlacementViewTests.Move(window, target);
            SystemPlacementViewTests.Release(window, target);

            Assert.Equal(
                ["250407 (long board)", "310378 Open128", "310378 (C128 & C128D)", "250477 (C128DCR)", "Issue 4B"],
                SystemPlacementViewTests.Boards(view));

            await view.SaveAsync();

            Assert.Equal(C64, Assert.Single(sent).AfterExcelDataFile);

            window.Close();
        });
    }

    // The systems already listed are CRT's order, not the maintainer's to rearrange here.
    [Fact]
    public void A_listed_system_does_not_move()
    {
        UiTest.Run(() =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());
            List<string> before = SystemPlacementViewTests.Boards(view);

            Point first = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 0));
            Point last = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 4));

            SystemPlacementViewTests.Press(window, first);
            SystemPlacementViewTests.Move(window, last);
            SystemPlacementViewTests.Release(window, last);

            Assert.Equal(before, SystemPlacementViewTests.Boards(view));
            Assert.DoesNotContain(view.Rows, row => row.IsDropPlaceholder);

            window.Close();
        });
    }

    // ###########################################################################################
    // A maintainer of ANOTHER system sees where it would go (every maintainer sees every system),
    // but cannot move it, rename it or save it - the server would refuse, so the screen says so
    // instead of offering it.
    // ###########################################################################################
    [Fact]
    public void An_account_that_cannot_place_it_sees_the_list_but_cannot_move_rename_or_save()
    {
        UiTest.Run(() =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(
                SystemPlacementViewTests.Listing(SystemPlacementViewTests.Entry(canPlace: false)));

            Assert.True(view.IsVisible);
            Assert.False(view.IsSaveEnabledForTests);
            Assert.False(view.AreNamesEditableForTests);
            Assert.EndsWith(SystemPlacementDisplay.CannotPlaceMessage, view.TextOfForTests("SavedStateText"), StringComparison.Ordinal);

            Point panel = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 3));
            Point top = SystemPlacementViewTests.CentreOf(window, SystemPlacementViewTests.RowContainer(view, 0));

            SystemPlacementViewTests.Press(window, panel);
            SystemPlacementViewTests.Move(window, top);
            SystemPlacementViewTests.Release(window, top);

            Assert.Equal("310378 Open128", SystemPlacementViewTests.Boards(view)[3]);

            // No grip: nothing says it moves.
            Assert.False(SystemPlacementViewTests.RowContainer(view, 3).GetVisualDescendants().OfType<TextBlock>()
                .Single(block => block.Classes.Contains("PlacementDragGrip")).IsVisible);

            window.Close();
        });
    }

    // The panel follows the names as they are typed, and names another system is listed under are
    // refused before Save - in the server's own words.
    [Fact]
    public void Typed_names_follow_into_the_panel_and_a_taken_pair_turns_Save_off_saying_why()
    {
        UiTest.Run(() =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());

            view.SetNamesForTests("Commodore 128", "250477 (C128DCR)", string.Empty);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("250477 (C128DCR)", view.Rows[3].BoardName);
            Assert.False(view.IsSaveEnabledForTests);
            Assert.Contains("is already in the drop-down lists, for Commodore/C128/250477", view.TextOfForTests("MessageText"));

            view.SetNamesForTests("Commodore 128D", "Open128", "Open-source replica.");
            Dispatcher.UIThread.RunJobs();

            Assert.True(view.IsSaveEnabledForTests);
            Assert.Equal(string.Empty, view.TextOfForTests("MessageText"));
            Assert.Equal("Commodore 128D is hardware 3 of 4 in CRT, and Open128 its board 1 of 1.", view.TextOfForTests("WhereInCrtText"));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE MINUTE CHECK NEVER TAKES AN UNSAVED PLACEMENT AWAY. *** The Systems screen re-reads the
    // lists every minute and hands them here. With the panel moved and not saved, it stays where the
    // maintainer put it; once saved, the next check shows the saved placement. Fails against a Show
    // that rebuilds every time.
    // ###########################################################################################
    [Fact]
    public async Task A_check_leaves_an_unsaved_placement_alone_and_shows_it_once_saved()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());
            SystemPlacementViewTests.AnswerSaves(view);

            view.SetNamesForTests("Commodore 128", "Open128", string.Empty);
            Dispatcher.UIThread.RunJobs();

            view.Show(SystemPlacementViewTests.ChangedListing(), Open128);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Open128", view.Rows[3].BoardName);
            Assert.Equal(5, view.Rows.Count);

            await view.SaveAsync();

            // What the server now holds - the next check shows it as saved.
            view.Show(
                SystemPlacementViewTests.Listing(SystemPlacementViewTests.Entry(new SystemPlacement("Commodore 128", "Open128", string.Empty, C128Dcr))),
                Open128);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Its place is saved.", view.TextOfForTests("SavedStateText"));
            Assert.Equal("Open128", view.Rows[3].BoardName);

            window.Close();
        });
    }

    // An unchanged answer does not rebuild the panel - the rows are the same objects.
    [Fact]
    public void An_unchanged_listing_does_not_rebuild_the_panel()
    {
        UiTest.Run(() =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());
            PlacementListRow first = view.Rows[0];

            view.Show(SystemPlacementViewTests.Listing(listed: SystemPlacementViewTests.Listed.ToList()), Open128);

            Assert.Same(first, view.Rows[0]);

            window.Close();
        });
    }

    [Fact]
    public async Task A_refused_save_keeps_the_placement_and_shows_the_servers_words()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, SystemPlacementView view) = SystemPlacementViewTests.Shown(SystemPlacementViewTests.Listing());

            view.SaveOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<SetPlacementAnswer>.Failed(ReviewApiFailure.Conflict, "The lists changed. Place it again."));

            view.SetNamesForTests("Commodore 128", "Open128", string.Empty);
            Dispatcher.UIThread.RunJobs();

            await view.SaveAsync();

            Assert.Equal("The lists changed. Place it again.", view.TextOfForTests("MessageText"));

            // Still unsaved: a check does not replace it.
            view.Show(SystemPlacementViewTests.ChangedListing(), Open128);
            Assert.Equal("Open128", view.Rows[3].BoardName);
            Assert.True(view.IsSaveEnabledForTests);

            window.Close();
        });
    }

    // A system the lists carry needs no placing - and no listing known is no panel either, never
    // "not in the lists" on the strength of a request that failed.
    [Fact]
    public void A_listed_system_or_an_unknown_listing_shows_no_panel()
    {
        UiTest.Run(() =>
        {
            var view = new SystemPlacementView();

            view.Show(SystemPlacementViewTests.Listing(), "Commodore/C64/250407");
            Assert.False(view.IsVisible);

            view.Show(null, Open128);
            Assert.False(view.IsVisible);

            view.Show(SystemPlacementViewTests.Listing(), Open128);
            Assert.True(view.IsVisible);

            view.Show(SystemPlacementViewTests.Listing(), "Commodore/C64/250407");
            Assert.False(view.IsVisible);
            Assert.Empty(view.Rows);
        });
    }
}
