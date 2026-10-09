using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers MaintainerModes - what the badges on the four screen buttons count (owner request,
// 2026-09-27: "should show a badge in button, how many systems need your attention").
//
// BOARDS, not submissions, and only those that wait for THIS account - the server's answer each
// time. A board waiting only for the other approver is not this account's to act on, and a badge
// that counted it would never go away.
// ###########################################################################################
public sealed class MaintainerModesTests
{
    private static ReviewQueueRow Row(long id, string board, bool? awaitsYou = true) =>
        new(id, board, "pending", "summary", "c@example.com", null, false, false, awaitsYou);

    private static ProductionBoardRow Beta(string board, bool? awaitsYou = true) =>
        new(board, "Commodore", "C64", "250407", null, "hash", null, null, awaitsYou);

    // Two submissions to one board are one board to look at.
    [Fact]
    public void Review_counts_the_boards_waiting_for_you_not_the_submissions()
    {
        Assert.Equal(2, MaintainerModes.ReviewAttention(
        [
            MaintainerModesTests.Row(1, "Commodore/C64/250407"),
            MaintainerModesTests.Row(2, "Commodore/C64/250407"),
            MaintainerModesTests.Row(3, "Commodore/C128/310378")
        ]));
    }

    // ###########################################################################################
    // A board whose only submission is with the other approver is not counted; one with a second
    // submission that IS yours still is. A row the server said nothing about counts as yours, the
    // way the queue shows it (undimmed).
    // ###########################################################################################
    [Fact]
    public void Review_leaves_out_a_board_that_waits_only_for_the_other_approver()
    {
        Assert.Equal(2, MaintainerModes.ReviewAttention(
        [
            MaintainerModesTests.Row(1, "Commodore/C64/250407", awaitsYou: false),
            MaintainerModesTests.Row(2, "Commodore/C128/310378", awaitsYou: false),
            MaintainerModesTests.Row(3, "Commodore/C128/310378", awaitsYou: true),
            MaintainerModesTests.Row(4, "Commodore/VIC-20/250403", awaitsYou: null)
        ]));

        Assert.Equal(0, MaintainerModes.ReviewAttention([]));
        Assert.Equal(0, MaintainerModes.ReviewAttention(null));
    }

    [Fact]
    public void Beta_counts_the_boards_that_wait_for_you()
    {
        Assert.Equal(2, MaintainerModes.BetaAttention(
        [
            MaintainerModesTests.Beta("Commodore/C64/250407", awaitsYou: true),
            MaintainerModesTests.Beta("Commodore/C128/310378", awaitsYou: false),
            MaintainerModesTests.Beta("Commodore/VIC-20/250403", awaitsYou: null)
        ]));

        Assert.Equal(0, MaintainerModes.BetaAttention(null));
    }

    // Nothing waiting is no badge at all; a backlog past 99 reads "99+" rather than pushing the
    // other buttons out of the row.
    [Fact]
    public void An_attention_badge_is_hidden_at_zero_and_capped_past_ninety_nine()
    {
        Assert.Null(MaintainerModes.AttentionBadge(0));
        Assert.Null(MaintainerModes.AttentionBadge(-1));
        Assert.Equal("1", MaintainerModes.AttentionBadge(1));
        Assert.Equal("99", MaintainerModes.AttentionBadge(99));
        Assert.Equal("99+", MaintainerModes.AttentionBadge(100));
    }

    // The Boards count is shown once the list has been read - zero included - and not before, so
    // a failed request never reads as "0 boards".
    [Fact]
    public void The_boards_count_shows_once_known_and_never_before()
    {
        Assert.Null(MaintainerModes.CountBadge(null));
        Assert.Equal("0", MaintainerModes.CountBadge(0));
        Assert.Equal("42", MaintainerModes.CountBadge(42));
    }

    // ###########################################################################################
    // *** THE TAB'S OWN BADGE IS THE TWO BUTTONS ADDED UP (owner request, 2026-09-30). *** A board
    // with a submission queued AND an earlier one in BETA waits for you twice, and each button
    // counts it - so the tab must say 3 over a 2 and a 1, or the maintainer opens it expecting two
    // things and finds three.
    // ###########################################################################################
    [Fact]
    public void The_tabs_badge_adds_up_the_review_and_BETA_buttons()
    {
        ReviewQueueRow[] queue =
        [
            MaintainerModesTests.Row(1, "Commodore/C64/250407"),
            MaintainerModesTests.Row(2, "Commodore/C128/310378")
        ];

        ProductionBoardRow[] beta = [MaintainerModesTests.Beta("Commodore/C64/250407")];

        Assert.Equal(
            MaintainerModes.ReviewAttention(queue) + MaintainerModes.BetaAttention(beta),
            MaintainerModes.TabAttention(queue, beta));

        Assert.Equal(3, MaintainerModes.TabAttention(queue, beta));
    }

    // "Until it is fully processed": nothing waiting for THIS account - including a board that
    // waits only for the other approver, in either list - is no badge at all.
    [Fact]
    public void The_tabs_badge_is_gone_when_nothing_waits_for_you()
    {
        Assert.Equal(0, MaintainerModes.TabAttention(
            [MaintainerModesTests.Row(1, "Commodore/C64/250407", awaitsYou: false)],
            [MaintainerModesTests.Beta("Commodore/C128/310378", awaitsYou: false)]));

        Assert.Equal(0, MaintainerModes.TabAttention(null, null));
        Assert.Null(MaintainerModes.AttentionBadge(0));
    }

    // ###########################################################################################
    // "SHOW ONLY THE MAINTAINER TAB WHEN I HAVE OUTSTANDING WORK" (owner request, 2026-10-01).
    //
    // The ordinary case first: with the new setting off, the tab follows "Enable Maintainer tab"
    // exactly as it always did - the badge decides nothing.
    // ###########################################################################################
    [Fact]
    public void With_the_setting_off_the_tab_follows_Enable_Maintainer_tab_alone()
    {
        Assert.True(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: false, badgeKnown: true, attention: 0, selected: false, unsavedEdits: false));

        Assert.False(MaintainerModes.TabIsShown(
            enabled: false, onlyWhenWaiting: false, badgeKnown: true, attention: 5, selected: false, unsavedEdits: false));

        // The new setting NARROWS the old one; it can never bring the tab back on its own.
        Assert.False(MaintainerModes.TabIsShown(
            enabled: false, onlyWhenWaiting: true, badgeKnown: true, attention: 5, selected: true, unsavedEdits: false));
    }

    // With it on, the badge's number is the condition: work waiting shows the tab, none hides it.
    [Fact]
    public void With_the_setting_on_the_tab_appears_only_while_something_waits()
    {
        Assert.True(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: true, badgeKnown: true, attention: 1, selected: false, unsavedEdits: false));

        Assert.False(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: true, badgeKnown: true, attention: 0, selected: false, unsavedEdits: false));
    }

    // ###########################################################################################
    // *** A TAB THE MAINTAINER IS STANDING ON IS NEVER TAKEN AWAY (owner decision, 2026-10-01). ***
    // Approving the last queued submission drops the badge to zero at that instant. Without this
    // the tab, the submission's table and any unsaved edits in it would vanish mid-review - so
    // `selected` holds it open, and it goes on the next tab switch instead.
    // ###########################################################################################
    [Fact]
    public void The_tab_being_looked_at_is_never_hidden_under_the_maintainer()
    {
        Assert.True(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: true, badgeKnown: true, attention: 0, selected: true, unsavedEdits: false));

        // ...and goes once it is left, the badge still being empty.
        Assert.False(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: true, badgeKnown: true, attention: 0, selected: false, unsavedEdits: false));
    }

    // ###########################################################################################
    // *** NOR IS A TAB HOLDING UNSAVED TABLE EDITS (code review, 2026-10-01). *** The edited
    // submission may wait only for the other approver, so it counts nothing; switching away
    // without saving then let the next minute check hide the tab with the edits inside it.
    // ###########################################################################################
    [Fact]
    public void A_tab_holding_unsaved_table_edits_is_never_hidden()
    {
        Assert.True(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: true, badgeKnown: true, attention: 0, selected: false, unsavedEdits: true));

        // Turning the tab off is still a deliberate act, and still wins.
        Assert.False(MaintainerModes.TabIsShown(
            enabled: false, onlyWhenWaiting: true, badgeKnown: true, attention: 0, selected: false, unsavedEdits: true));
    }

    // ###########################################################################################
    // *** AN UNREAD BADGE IS NOT "NO WORK" (code review, 2026-10-01). *** Zero is also what the badge
    // shows before the lists are read, after a 401 empties it, and while the server cannot be
    // reached. Hiding the tab then would remove the only way back to the sign-in screen, or make the
    // tab vanish on a launch whose first read failed - with only the Configuration tab to bring it
    // back. TabMaintainer.BadgeKnown is false until somebody is signed in and both lists were read.
    // ###########################################################################################
    [Fact]
    public void A_badge_that_has_not_been_read_never_hides_the_tab()
    {
        Assert.True(MaintainerModes.TabIsShown(
            enabled: true, onlyWhenWaiting: true, badgeKnown: false, attention: 0, selected: false, unsavedEdits: false));
    }

    // ###########################################################################################
    // Which entry "Contributor Submissions" and "Beta > Prod" open on (owner request, 2026-09-30):
    // the one looked at last while it is still listed, else the first as the list shows it.
    // ###########################################################################################
    [Fact]
    public void A_screen_opens_on_the_entry_looked_at_last_or_else_the_first()
    {
        Assert.Equal(9, MaintainerModes.EntryToOpen([7L, 3L, 9L], 9));
        Assert.Equal(7, MaintainerModes.EntryToOpen([7L, 3L, 9L], 12345));
        Assert.Equal(7, MaintainerModes.EntryToOpen([7L, 3L, 9L], null));
        Assert.Null(MaintainerModes.EntryToOpen(Array.Empty<long>(), 9));

        Assert.Equal("Commodore/C128/310378", MaintainerModes.EntryToOpen(["Commodore/C64/250407", "Commodore/C128/310378"], "Commodore/C128/310378"));
        Assert.Equal("Commodore/C64/250407", MaintainerModes.EntryToOpen(["Commodore/C64/250407", "Commodore/C128/310378"], "commodore/c128/310378"));
        Assert.Null(MaintainerModes.EntryToOpen(Array.Empty<string>(), null));
    }

    // ###########################################################################################
    // The screen the tab opens on (owner request, 2026-10-04: "if there is no queue awaiting, when
    // opening the "Maintainer" tab, then go to "Systems" and show the last selected system"):
    // Boards when nothing waits for this account in either queue - a queue screen gives way,
    // unless something is open on it; Boards and Account, chosen, stay.
    // ###########################################################################################
    [Theory]
    [InlineData(MaintainerMode.Review, 0, false, MaintainerMode.Boards)]
    [InlineData(MaintainerMode.Beta, 0, false, MaintainerMode.Boards)]
    [InlineData(MaintainerMode.Review, 1, false, MaintainerMode.Review)]
    [InlineData(MaintainerMode.Beta, 2, false, MaintainerMode.Beta)]
    [InlineData(MaintainerMode.Review, 0, true, MaintainerMode.Review)]
    [InlineData(MaintainerMode.Beta, 0, true, MaintainerMode.Beta)]
    [InlineData(MaintainerMode.Boards, 3, false, MaintainerMode.Boards)]
    [InlineData(MaintainerMode.Boards, 0, false, MaintainerMode.Boards)]
    [InlineData(MaintainerMode.Account, 0, false, MaintainerMode.Account)]
    [InlineData(MaintainerMode.Account, 4, false, MaintainerMode.Account)]
    public void The_tab_opens_on_Boards_when_nothing_waits_and_nothing_is_open(
        MaintainerMode shown, int attention, bool somethingOpen, MaintainerMode expected)
    {
        Assert.Equal(expected, MaintainerModes.ScreenOnOpening(shown, attention, somethingOpen));
    }
}
