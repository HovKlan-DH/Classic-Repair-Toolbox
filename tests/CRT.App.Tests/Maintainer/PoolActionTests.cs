using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// PoolAction - one change to a system's maintainers on the Systems screen, and after a change that
// had no answer within the two-minute limit, whether the system as the server now describes it
// holds that change (2026-09-28: "it must be solid in validating if it did finish").
// ###########################################################################################
public sealed class PoolActionTests
{
    private const string System = "Commodore/C64/250407";

    private static SystemDetailAnswer Detail(long[] maintainers, params MaintainerInvitationEntry[] invitations) =>
        new(
            new SystemOverviewEntry(System, "Commodore", "C64", "250407", true, true, false, true, null, null, null, maintainers.Length),
            [.. maintainers.Select(id => new PoolMaintainerEntry(id, $"Account {id}", $"a{id}@example.com"))],
            [],
            [],
            invitations);

    private static MaintainerInvitationEntry Invitation(long id, string email) =>
        new(id, email, DateTimeOffset.UnixEpoch, DateTimeOffset.UnixEpoch.AddDays(7));

    [Fact]
    public void An_added_maintainer_has_landed_once_they_are_in_the_pool()
    {
        var add = new PoolAction(PoolActionKind.Add, System, AccountId: 7, Who: "Dennis");

        Assert.True(add.HasLanded(Detail([3, 7])));
        Assert.False(add.HasLanded(Detail([3])));
    }

    [Fact]
    public void A_removed_maintainer_has_landed_once_they_are_gone()
    {
        var remove = new PoolAction(PoolActionKind.Remove, System, AccountId: 7, Who: "Dennis");

        Assert.True(remove.HasLanded(Detail([3])));
        Assert.False(remove.HasLanded(Detail([3, 7])));
    }

    // An invitation has no id until it exists, so it is found by its address - in any case.
    [Fact]
    public void An_invitation_has_landed_once_one_is_open_for_that_address()
    {
        var invite = new PoolAction(PoolActionKind.Invite, System, Email: "new@example.com", Who: "new@example.com");

        Assert.True(invite.HasLanded(Detail([], Invitation(4, "NEW@example.com"))));
        Assert.False(invite.HasLanded(Detail([], Invitation(4, "other@example.com"))));
    }

    [Fact]
    public void A_withdrawal_has_landed_once_that_invitation_is_gone()
    {
        var withdraw = new PoolAction(PoolActionKind.Withdraw, System, InvitationId: 4, Who: "new@example.com");

        Assert.True(withdraw.HasLanded(Detail([], Invitation(5, "new@example.com"))));
        Assert.False(withdraw.HasLanded(Detail([], Invitation(4, "new@example.com"))));
    }

    [Fact]
    public void Each_change_names_the_person_while_it_waits_and_afterwards()
    {
        var add = new PoolAction(PoolActionKind.Add, System, AccountId: 7, Who: "Dennis");
        var withdraw = new PoolAction(PoolActionKind.Withdraw, System, InvitationId: 4, Who: "new@example.com");

        Assert.Equal($"Adding Dennis as a maintainer of {System}...", add.Waiting);
        Assert.Equal($"Dennis now maintains {System}", add.DoneClause);
        Assert.Equal($"Dennis does not maintain {System} yet", add.NotDoneClause);
        Assert.Equal("Withdrawing the invitation to new@example.com...", withdraw.Waiting);
        Assert.Equal("the invitation to new@example.com is still open", withdraw.NotDoneClause);
    }
}
