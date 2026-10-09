using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ONE CHANGE TO A BOARD'S MAINTAINERS, and how it reads (the Boards screen's pool section,
    // BoardDetailView.Maintainers.cs). Pure - wording, and whether the change is in a BoardDetailAnswer -
    // so it lives here rather than in the tab (code review, 2026-09-29: CLAUDE.md's "pure logic does
    // not belong in a tab"; it was declared in BoardDetailView.Maintainers.cs). Tested by PoolActionTests.
    // ###########################################################################################
    internal enum PoolActionKind
    {
        Add,
        Remove,
        Invite,
        Withdraw
    }

    // One change to a board's maintainers, as the administrator asked for it.
    // `Who` names the person in what the overlay and the after-timeout sentence say: the account's
    // name for Add and Remove, the address for Invite and Withdraw.
    internal sealed record PoolAction(
        PoolActionKind Kind,
        string BoardId,
        long AccountId = 0,
        string Email = "",
        long InvitationId = 0,
        string Who = "")
    {
        public string Waiting => this.Kind switch
        {
            PoolActionKind.Add => MaintainerWaitWording.AddingMaintainer(this.Who, this.BoardId),
            PoolActionKind.Remove => MaintainerWaitWording.RemovingMaintainer(this.Who, this.BoardId),
            PoolActionKind.Invite => MaintainerWaitWording.Inviting(this.Who, this.BoardId),
            _ => MaintainerWaitWording.WithdrawingInvitation(this.Who)
        };

        public string DoneClause => this.Kind switch
        {
            PoolActionKind.Add => $"{this.Who} now maintains {this.BoardId}",
            PoolActionKind.Remove => $"{this.Who} no longer maintains {this.BoardId}",
            PoolActionKind.Invite => $"{this.Who} is invited to maintain {this.BoardId}",
            _ => $"the invitation to {this.Who} is withdrawn"
        };

        public string NotDoneClause => this.Kind switch
        {
            PoolActionKind.Add => $"{this.Who} does not maintain {this.BoardId} yet",
            PoolActionKind.Remove => $"{this.Who} still maintains {this.BoardId}",
            PoolActionKind.Invite => $"{this.Who} is not invited yet",
            _ => $"the invitation to {this.Who} is still open"
        };

        // ###########################################################################################
        // Whether this change is in the board as the server now describes it - read after the
        // two-minute limit passed with no answer. An invitation is found by its address (it has no
        // id until it exists), a withdrawal by the invitation's id.
        // ###########################################################################################
        public bool HasLanded(BoardDetailAnswer detail)
        {
            ArgumentNullException.ThrowIfNull(detail);

            bool maintains = detail.Maintainers.Any(maintainer => maintainer.AccountId == this.AccountId);
            IReadOnlyList<MaintainerInvitationEntry> invitations = detail.Invitations ?? [];

            return this.Kind switch
            {
                PoolActionKind.Add => maintains,
                PoolActionKind.Remove => !maintains,
                PoolActionKind.Invite => invitations.Any(invitation => string.Equals(invitation.Email, this.Email, StringComparison.OrdinalIgnoreCase)),
                _ => invitations.All(invitation => invitation.Id != this.InvitationId)
            };
        }
    }
}
