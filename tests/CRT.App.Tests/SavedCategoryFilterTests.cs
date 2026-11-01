using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SavedCategoryFilter - which categories the filter beside the schematics starts with ticked, once
// somebody has narrowed it on a board. A category it has never seen - renamed, typed for one component,
// edited in the table, or come with a data update - is ticked: left out, every component under it
// dropped out of the list with nothing saying why (code review, 2026-11-01).
// ###########################################################################################
public sealed class SavedCategoryFilterTests
{
    private static readonly string[] Saved = ["IC"];
    private static readonly string[] Seen = ["IC", "Capacitor"];

    [Theory]
    [InlineData("IC", true)]
    [InlineData("ic", true)]
    [InlineData("Capacitor", false)]
    [InlineData("capacitor", false)]
    [InlineData("Chips", true)]
    public void A_category_is_ticked_when_chosen_or_never_seen(string category, bool ticked)
    {
        Assert.Equal(ticked, SavedCategoryFilter.IsTicked(category, Saved, Seen));
    }

    // Saved before what was seen was kept, nothing can be told apart: the list alone decides, as it did.
    [Theory]
    [InlineData("IC", true)]
    [InlineData("Capacitor", false)]
    [InlineData("Chips", false)]
    public void A_filter_saved_before_this_was_kept_goes_by_its_list(string category, bool ticked)
    {
        Assert.Equal(ticked, SavedCategoryFilter.IsTicked(category, Saved, null));
    }
}
