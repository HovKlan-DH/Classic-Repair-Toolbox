using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// ApprovalGate - why Approve is off for a submission this account could otherwise approve
// (2026-09-27): a system with an earlier submission still in BETA (one submission in BETA per
// system), or a new system with no place in the drop-down lists yet.
// ###########################################################################################
public sealed class ApprovalGateTests
{
    private const string C64 = "Commodore/C64/250407";

    private static ProductionSystemRow Beta(string system) =>
        new(system, "Commodore", "C64", "250407", "2026-September-27", "hash", null, null, true);

    private static SystemListingAnswer Unplaced(string system) =>
        new(true, [], [new UnlistedSystemEntry(system, "Commodore", "C64", "250407", false, true, null, new SystemPlacement("C64", "250407", string.Empty, null))]);

    // *** THE SERVER'S OWN SENTENCE *** - CRT.Data's, which the approval refusal uses too.
    [Fact]
    public void A_system_waiting_in_BETA_blocks_approval_in_the_servers_words()
    {
        Assert.Equal(OneSubmissionInBeta.BusyMessage(C64), ApprovalGate.Blocked(C64, null, [ApprovalGateTests.Beta(C64)]));
        Assert.Contains("takes one submission into BETA at a time", ApprovalGate.Blocked(C64, null, [ApprovalGateTests.Beta(C64)]));
    }

    [Fact]
    public void A_new_system_with_no_place_in_the_lists_blocks_approval()
    {
        Assert.Contains("no place in CRT's drop-down lists yet", ApprovalGate.Blocked(C64, ApprovalGateTests.Unplaced(C64), []));
    }

    // Another system in BETA, or a placed or listed one, is no reason.
    [Fact]
    public void Nothing_blocks_a_listed_system_with_nothing_in_BETA()
    {
        Assert.Null(ApprovalGate.Blocked(C64, new SystemListingAnswer(true, [], []), [ApprovalGateTests.Beta("Commodore/C128/310378")]));
        Assert.Null(ApprovalGate.Blocked(C64, null, null));
        Assert.Null(ApprovalGate.Blocked(null, ApprovalGateTests.Unplaced(C64), [ApprovalGateTests.Beta(C64)]));
    }
}
