using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// InPlaceList - changing a shown list without replacing the items still on screen (2026-10-10):
// a ComboBox whose items are replaced lets go of its selection (the "Schematic images" window's CAD
// name drop-downs), and a row replaced between a click's press and release loses the click (the
// Drafts tab's list). Pure list work, tested on plain lists - BoardFilesWindowTests and
// TabDraftsTests cover the two lists on screen.
// ###########################################################################################
public sealed class InPlaceListTests
{
    // An item with an identity of its own, like a drop-down entry or a list row.
    private sealed class Item(string name)
    {
        public string Name { get; } = name;
    }

    private static void Sync(List<Item> items, string[] wanted, Item? keep = null) =>
        InPlaceList.Sync(items, wanted, (item, name) => item.Name == name, name => new Item(name), keep);

    private static List<string> Names(List<Item> items) => items.Select(item => item.Name).ToList();

    [Fact]
    public void Entries_already_there_keep_their_objects_and_new_ones_are_inserted_in_order()
    {
        Item none = new("");
        Item reset = new("Reset");
        var items = new List<Item> { none, reset };

        Sync(items, ["", "Clock", "Reset", "VIC"], keep: reset);

        Assert.Equal(["", "Clock", "Reset", "VIC"], Names(items));
        Assert.Same(none, items[0]);
        Assert.Same(reset, items[2]);
    }

    [Fact]
    public void An_entry_no_longer_wanted_is_removed_and_the_others_stay_where_they_are()
    {
        Item reset = new("Reset");
        Item vic = new("VIC");
        var items = new List<Item> { new(""), new("Clock"), reset, vic };

        Sync(items, ["", "Reset", "VIC"], keep: reset);

        Assert.Equal(["", "Reset", "VIC"], Names(items));
        Assert.Same(reset, items[1]);
        Assert.Same(vic, items[2]);
    }

    // ###########################################################################################
    // *** THE ITEM ON SCREEN IS NEVER REMOVED TO REACH ONE BELOW IT. *** "B" is wanted before the
    // selected "A": rather than removing "A" on the way to the "B" further down and making "A" again
    // (which a ComboBox reads as its selection gone), a new "B" goes in above it and the old one
    // drops off the end.
    // ###########################################################################################
    [Fact]
    public void The_kept_item_is_not_removed_to_reach_an_entry_below_it()
    {
        Item a = new("A");
        var items = new List<Item> { a, new("B") };

        Sync(items, ["B", "A"], keep: a);

        Assert.Equal(["B", "A"], Names(items));
        Assert.Same(a, items[1]);
    }

    [Fact]
    public void Wanting_nothing_empties_the_list_and_wanting_from_nothing_fills_it()
    {
        var items = new List<Item> { new("A"), new("B") };

        Sync(items, []);
        Assert.Empty(items);

        Sync(items, ["X", "Y"]);
        Assert.Equal(["X", "Y"], Names(items));
    }

    // ------------------------------------------------------------------ KeepUnchanged

    [Fact]
    public void Unchanged_rows_are_kept_changed_ones_replaced_and_the_end_follows_the_fresh_list()
    {
        Item a = new("A");
        Item b = new("B");
        Item c = new("C");
        var rows = new List<Item> { a, b, c };

        Item freshA = new("A");
        Item freshB = new("B (changed)");

        InPlaceList.KeepUnchanged(rows, [freshA, freshB], (row, fresh) => row.Name == fresh.Name);

        Assert.Same(a, rows[0]);
        Assert.Same(freshB, rows[1]);
        Assert.Equal(2, rows.Count);

        Item freshC = new("C");
        InPlaceList.KeepUnchanged(rows, [freshA, freshB, freshC], (row, fresh) => row.Name == fresh.Name);

        Assert.Same(a, rows[0]);
        Assert.Same(freshB, rows[1]);
        Assert.Same(freshC, rows[2]);
    }
}
