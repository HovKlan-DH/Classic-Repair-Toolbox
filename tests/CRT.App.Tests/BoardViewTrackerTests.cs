using System;
using Handlers.DataHandling;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardViewTracker - when a board counts as viewed (owner request, 2026-09-27: "catch every time a
// user selects a board ... count only when viewed for +10 seconds").
//
// Every time a board comes on screen and stays ten seconds is one view; a board left sooner is not;
// the same board re-read (not chosen again) is not a new view; and a view counts once.
// ###########################################################################################
public sealed class BoardViewTrackerTests
{
    private static readonly DateTimeOffset Start = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

    private const string C64 = "Commodore/C64/250407";
    private const string Vic = "Commodore/VIC-20/250403";

    private static DateTimeOffset At(double seconds) => BoardViewTrackerTests.Start.AddSeconds(seconds);

    [Fact]
    public void A_board_counts_once_it_has_been_on_screen_ten_seconds()
    {
        var tracker = new BoardViewTracker();

        Assert.True(tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0)));

        Assert.Null(tracker.TakeDue(BoardViewTrackerTests.At(9.9)));
        Assert.Equal(BoardViewTrackerTests.C64, tracker.TakeDue(BoardViewTrackerTests.At(10)));
    }

    // However long it then stays, it is one view.
    [Fact]
    public void A_view_counts_once_however_long_the_board_stays()
    {
        var tracker = new BoardViewTracker();
        tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0));

        Assert.Equal(BoardViewTrackerTests.C64, tracker.TakeDue(BoardViewTrackerTests.At(10)));
        Assert.Null(tracker.TakeDue(BoardViewTrackerTests.At(600)));
        Assert.Null(tracker.Remaining(BoardViewTrackerTests.At(600)));
    }

    // Flicking through the drop-down counts nothing: a board left within ten seconds never counts.
    [Fact]
    public void A_board_left_within_ten_seconds_is_not_counted()
    {
        var tracker = new BoardViewTracker();

        tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0));
        Assert.True(tracker.Show(BoardViewTrackerTests.Vic, BoardViewTrackerTests.At(4)));

        Assert.Null(tracker.TakeDue(BoardViewTrackerTests.At(10)));
        Assert.Equal(BoardViewTrackerTests.Vic, tracker.TakeDue(BoardViewTrackerTests.At(14)));
    }

    // ###########################################################################################
    // *** THE SAME BOARD SHOWN AGAIN IS NOT A NEW VIEW. *** Hiding a schematic, or a draft saved,
    // re-reads the board on screen - not the user choosing it. It neither restarts a waiting view
    // nor counts a finished one again.
    // ###########################################################################################
    [Fact]
    public void The_same_board_re_read_is_not_a_new_view()
    {
        var tracker = new BoardViewTracker();

        tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0));
        Assert.False(tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(8)));

        // Not restarted: it still counts at ten seconds from the first showing.
        Assert.Equal(BoardViewTrackerTests.C64, tracker.TakeDue(BoardViewTrackerTests.At(10)));

        Assert.False(tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(30)));
        Assert.Null(tracker.TakeDue(BoardViewTrackerTests.At(45)));
    }

    // "Every time a user selects a board": away and back is a second view.
    [Fact]
    public void Going_to_another_board_and_back_counts_the_first_board_again()
    {
        var tracker = new BoardViewTracker();

        tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0));
        Assert.Equal(BoardViewTrackerTests.C64, tracker.TakeDue(BoardViewTrackerTests.At(10)));

        tracker.Show(BoardViewTrackerTests.Vic, BoardViewTrackerTests.At(20));
        Assert.Equal(BoardViewTrackerTests.Vic, tracker.TakeDue(BoardViewTrackerTests.At(30)));

        Assert.True(tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(40)));
        Assert.Equal(BoardViewTrackerTests.C64, tracker.TakeDue(BoardViewTrackerTests.At(50)));
    }

    // No board (none selected, one that failed to load, or one not counted) waits for nothing, and
    // ends the view waiting for the board before it.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void No_board_on_screen_counts_nothing_and_ends_the_waiting_view(string? none)
    {
        var tracker = new BoardViewTracker();

        tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0));
        Assert.True(tracker.Show(none, BoardViewTrackerTests.At(5)));

        Assert.Null(tracker.Shown);
        Assert.Null(tracker.Remaining(BoardViewTrackerTests.At(20)));
        Assert.Null(tracker.TakeDue(BoardViewTrackerTests.At(20)));
    }

    // What the timer is set to: the rest of the ten seconds, and zero once they are up.
    [Fact]
    public void Remaining_is_the_rest_of_the_ten_seconds()
    {
        var tracker = new BoardViewTracker();
        tracker.Show(BoardViewTrackerTests.C64, BoardViewTrackerTests.At(0));

        Assert.Equal(BoardViewRules.MinimumTimeOnScreen, tracker.Remaining(BoardViewTrackerTests.At(0)));
        Assert.Equal(TimeSpan.FromSeconds(3), tracker.Remaining(BoardViewTrackerTests.At(7)));
        Assert.Equal(TimeSpan.Zero, tracker.Remaining(BoardViewTrackerTests.At(25)));
    }
}
