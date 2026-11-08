using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The id a new row of a record's own list takes - RecordIds, the one rule for a worklog's and a device's lists (code
// review, 2026-11-08: ten copies of it).
// ###########################################################################################
public sealed class RecordIdsTests
{
    [Fact]
    public void A_new_row_takes_one_past_the_highest_id_and_the_first_takes_1()
    {
        Assert.Equal(1, RecordIds.Next([]));
        Assert.Equal(2, RecordIds.Next([1]));
        Assert.Equal(8, RecordIds.Next([3, 7, 2]));

        // Within a list only: an id freed by a delete at the top comes back.
        Assert.Equal(3, RecordIds.Next([1, 2]));
    }
}
