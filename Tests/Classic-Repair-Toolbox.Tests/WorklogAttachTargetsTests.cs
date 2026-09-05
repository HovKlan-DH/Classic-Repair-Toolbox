using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers the ranking behind the "attach this oscilloscope capture to a worklog" picker.
//
// The whole value of this class is that the FIRST row is the one the user wants, since the dialog
// preselects it and a correct guess turns the flow into a single Attach click. So most of these
// tests assert on position, not merely on presence - a list containing the right entry somewhere is
// not the same as a list offering it first.
//
// Pure logic with no statics of its own, so this class joins no xUnit collection.
public sealed class WorklogAttachTargetsTests
{
    private static WorklogEntryRecord Entry(
        int id,
        string title = "",
        string state = "Open",
        params string[] componentLabels) =>
        new()
        {
            Id = id,
            Title = title,
            State = state,
            ComponentLabels = componentLabels.ToList()
        };

    // The headline rule: someone probing U8 while working a fault on U8 gets that fault offered
    // first, even though a newer entry exists that would otherwise win on id.
    [Fact]
    public void An_entry_scoping_the_measured_component_is_offered_first()
    {
        var entries = new List<WorklogEntryRecord>
        {
            Entry(1, "Video fault", "Open", "U8"),
            Entry(2, "Keyboard fault", "Open", "U1")
        };

        var ranked = WorklogAttachTargets.Rank(entries, "U8");

        Assert.Equal(1, ranked[0].Entry.Id);
        Assert.True(ranked[0].IsComponentMatch);
        Assert.False(ranked[1].IsComponentMatch);
    }

    // Component matching follows every other BoardLabel comparison in the app: case-insensitive and
    // trimmed, so a label stored as "u8 " still matches a capture taken on "U8".
    [Theory]
    [InlineData("u8")]
    [InlineData("U8")]
    [InlineData("  U8  ")]
    public void Component_matching_ignores_case_and_surrounding_space(string componentLabel)
    {
        var entries = new List<WorklogEntryRecord> { Entry(1, "Video fault", "Open", " u8 ") };

        var ranked = WorklogAttachTargets.Rank(entries, componentLabel);

        Assert.True(ranked[0].IsComponentMatch);
    }

    // With no component match to lead on, the list is in plain counting order. This renders as a
    // dropdown, and a LIST has to look ordered or it looks broken - an earlier version sorted
    // newest-first inside an open-before-closed band and produced "#2, #4, #3, #1" on screen, which
    // reads as no order at all. Entry ids are also what the board pills show, so this is the one
    // ordering the user can already follow.
    [Fact]
    public void Without_a_component_match_the_list_is_in_plain_id_order()
    {
        var entries = new List<WorklogEntryRecord>
        {
            Entry(1, "Older"),
            Entry(3, "Newest"),
            Entry(2, "Middle")
        };

        var ranked = WorklogAttachTargets.Rank(entries, "U8");

        Assert.Equal(new[] { 1, 2, 3 }, ranked.Select(target => target.Entry.Id).ToArray());
    }

    // Closed entries are KEPT - a board that comes back is re-measured against the entry describing
    // the original repair - and are NOT sorted apart from the open ones: open/closed is not a sort
    // level, because a second invisible criterion is what made the list look unordered. Both values
    // already carry an always-visible pill in the editor, so they read by eye rather than by
    // position.
    [Fact]
    public void Closed_entries_are_not_hidden_and_do_not_break_the_id_order()
    {
        var entries = new List<WorklogEntryRecord>
        {
            Entry(5, "Finished repair", "Closed"),
            Entry(2, "Still open", "Open"),
            Entry(3, "Also finished", "Closed")
        };

        var ranked = WorklogAttachTargets.Rank(entries, "U8");

        Assert.Equal(new[] { 2, 3, 5 }, ranked.Select(target => target.Entry.Id).ToArray());
    }

    // The component match is the ONE surviving sort level above id, because it pays for itself - it
    // puts the right answer in the preselected slot - and the dialog says out loud that it is doing
    // it. A CLOSED entry scoping the measured component still leads an unrelated open one: the
    // "board came back with the same fault" case, which is precisely when someone re-probes.
    [Fact]
    public void A_component_match_outranks_an_unrelated_open_entry_even_when_closed()
    {
        var entries = new List<WorklogEntryRecord>
        {
            Entry(1, "Replaced U8", "Closed", "U8"),
            Entry(2, "Unrelated", "Open", "U1")
        };

        var ranked = WorklogAttachTargets.Rank(entries, "U8");

        Assert.Equal(1, ranked[0].Entry.Id);
    }

    // A capture whose component cannot be resolved must still produce a usable list rather than
    // matching everything - a blank component means "no match", not "match all".
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void A_blank_component_matches_nothing_and_still_lists_in_id_order(string? componentLabel)
    {
        var entries = new List<WorklogEntryRecord>
        {
            Entry(1, "One", "Open", "U8"),
            Entry(2, "Two", "Open", "U1")
        };

        var ranked = WorklogAttachTargets.Rank(entries, componentLabel);

        Assert.All(ranked, target => Assert.False(target.IsComponentMatch));
        Assert.Equal(new[] { 1, 2 }, ranked.Select(target => target.Entry.Id).ToArray());
    }

    // The matched band is itself in id order, so the "listed first" group reads as a list too rather
    // than swapping the disorder from the whole dropdown into its top few rows.
    [Fact]
    public void Several_component_matches_stay_in_id_order_among_themselves()
    {
        var entries = new List<WorklogEntryRecord>
        {
            Entry(4, "Also U8", "Open", "U8"),
            Entry(2, "Unrelated", "Open", "U1"),
            Entry(1, "First U8", "Open", "U8")
        };

        var ranked = WorklogAttachTargets.Rank(entries, "U8");

        Assert.Equal(new[] { 1, 4, 2 }, ranked.Select(target => target.Entry.Id).ToArray());
    }

    // The caller reads this straight off GetEntries for a workbook that may have none yet, so an
    // empty or null list is a normal input rather than a failure.
    [Fact]
    public void No_entries_yields_no_targets()
    {
        Assert.Empty(WorklogAttachTargets.Rank(new List<WorklogEntryRecord>(), "U8"));
        Assert.Empty(WorklogAttachTargets.Rank(null, "U8"));
    }

    // ComponentLabels records what an entry SCOPES; CompletedComponentLabels records which of those
    // have been dealt with. An entry that scopes U8 is a candidate whether or not U8 is ticked off,
    // so the match must not consult the completed list.
    [Fact]
    public void A_completed_component_still_counts_as_a_match()
    {
        var entry = Entry(1, "Replaced it", "Open", "U8");
        entry.CompletedComponentLabels = new List<string> { "U8" };

        var ranked = WorklogAttachTargets.Rank(new List<WorklogEntryRecord> { entry }, "U8");

        Assert.True(ranked[0].IsComponentMatch);
    }

    [Fact]
    public void An_entry_label_reads_as_id_and_title()
    {
        Assert.Equal("#7 - U8 gives no video", WorklogAttachTargets.FormatLabel(Entry(7, "U8 gives no video")));
    }

    // Title is a plain string in entries.json, so a hand-edited or older-build record can carry an
    // empty one. A bare "#7 - " with nothing after it reads as a rendering fault, so the id stands
    // alone instead.
    [Fact]
    public void An_entry_with_no_title_is_labelled_by_its_id_alone()
    {
        Assert.Equal("#7", WorklogAttachTargets.FormatLabel(Entry(7, "   ")));
        Assert.Equal(string.Empty, WorklogAttachTargets.FormatLabel(null));
    }

    // The workbook line uses the same shape and the same blank-title fallback, so the two lines in
    // the dialog read as one family.
    [Fact]
    public void A_workbook_label_matches_the_entry_label_shape()
    {
        var workbook = new WorkbookRecord { Id = 3, Title = "Dave's C64" };

        Assert.Equal("#3 - Dave's C64", WorklogAttachTargets.FormatWorkbookLabel(workbook));
        Assert.Equal("#3", WorklogAttachTargets.FormatWorkbookLabel(new WorkbookRecord { Id = 3, Title = "" }));
        Assert.Equal(string.Empty, WorklogAttachTargets.FormatWorkbookLabel(null));
    }
}
