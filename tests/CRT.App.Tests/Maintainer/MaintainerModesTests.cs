using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers MaintainerModes - what the badges on the four screen buttons count (owner request,
// 2026-09-27: "should show a badge in button, how many systems need your attention").
//
// SYSTEMS, not submissions, and only those that wait for THIS account - the server's answer each
// time. A board waiting only for the other approver is not this account's to act on, and a badge
// that counted it would never go away.
// ###########################################################################################
public sealed class MaintainerModesTests
{
    private static ReviewQueueRow Row(long id, string system, bool? awaitsYou = true) =>
        new(id, system, "pending", "summary", "c@example.com", null, false, false, awaitsYou);

    private static ProductionSystemRow Beta(string system, bool? awaitsYou = true) =>
        new(system, "Commodore", "C64", "250407", null, "hash", null, null, awaitsYou);

    // Two submissions to one board are one system to look at.
    [Fact]
    public void Review_counts_the_systems_waiting_for_you_not_the_submissions()
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
    public void Review_leaves_out_a_system_that_waits_only_for_the_other_approver()
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
    public void Beta_counts_the_systems_that_wait_for_you()
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

    // The Systems count is shown once the list has been read - zero included - and not before, so
    // a failed request never reads as "0 systems".
    [Fact]
    public void The_systems_count_shows_once_known_and_never_before()
    {
        Assert.Null(MaintainerModes.CountBadge(null));
        Assert.Equal("0", MaintainerModes.CountBadge(0));
        Assert.Equal("42", MaintainerModes.CountBadge(42));
    }

    // The tooltip says what the screen is, and then what its badge counts.
    [Fact]
    public void A_buttons_tooltip_says_what_the_screen_is_and_what_its_badge_counts()
    {
        Assert.Equal(
            "Contributions waiting for review.\n3 systems with contributions waiting for you.",
            MaintainerModes.Tooltip(MaintainerMode.Review, 3));

        Assert.Equal(
            "Systems in BETA, waiting to be published to production or pushed back to the queue.\n1 system waiting for you.",
            MaintainerModes.Tooltip(MaintainerMode.Beta, 1));

        Assert.EndsWith("\nNothing is waiting for you.", MaintainerModes.Tooltip(MaintainerMode.Beta, 0), StringComparison.Ordinal);
        Assert.EndsWith("\n42 systems in all.", MaintainerModes.Tooltip(MaintainerMode.Systems, 42), StringComparison.Ordinal);

        // New systems waiting for a place in the drop-down lists come first on the Systems button,
        // since they are what its badge then counts (2026-09-27).
        Assert.Equal(
            "Every system - who maintains it, who has contributed and how that went.\n" +
            "2 new systems need a place in the drop-down lists.\n42 systems in all.",
            MaintainerModes.Tooltip(MaintainerMode.Systems, 42, needingPlace: 2));
        Assert.EndsWith(
            "\n1 new system needs a place in the drop-down lists.\n42 systems in all.",
            MaintainerModes.Tooltip(MaintainerMode.Systems, 42, needingPlace: 1),
            StringComparison.Ordinal);

        // Not known yet: the screen only.
        Assert.Equal("Contributions waiting for review.", MaintainerModes.Tooltip(MaintainerMode.Review, null));
        Assert.Equal("Who maintains which system, and files nothing uses.", MaintainerModes.Tooltip(MaintainerMode.Admin, null));
    }
}
