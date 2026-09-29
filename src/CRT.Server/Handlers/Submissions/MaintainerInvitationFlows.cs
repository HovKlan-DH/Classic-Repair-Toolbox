using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // INVITING A NEW MAINTAINER BY EMAIL (owner request, 2026-09-27: "I should be able to either
    // select an existing maintainer or invite a new maintainer via email").
    //
    // Nobody can make an account for themselves in either application - accounts were registered by
    // hand - so an invitation is also how a new maintainer's account comes to exist:
    //
    //   1. The ADMINISTRATOR names a system and an address with no account (InviteAsync). The
    //      server mails the address a one-time code and keeps only its hash (migration 0012).
    //   2. The person opens the Maintainer tab, chooses "I have an invitation", pastes the code and
    //      picks a name and a password (AcceptAsync). That creates the account VERIFIED - the code
    //      reaching the mailbox proves what the verification link would - and puts it in the pool
    //      of EVERY system that address has an open invitation to.
    //
    // *** AN ADDRESS THAT HAS AN ACCOUNT IS NOT INVITED. *** It is chosen from the list of accounts
    // instead (MaintainerAssignmentFlows.AddAsync), which also says why an account cannot be granted
    // - unverified, locked, administrator. Inviting it would make a second way into the same pool
    // with none of those checks.
    //
    // *** THE PASSWORD IS HASHED ONLY AFTER THE CODE CHECKS OUT *** - the order the password reset
    // keeps - so the unauthenticated acceptance route cannot be used to make the server run the
    // ~128 MiB Argon2 hasher without holding a real, open invitation.
    //
    // Pure, taking its stores and mailer as arguments, like every other flow: the endpoints are a rim.
    // ###########################################################################################
    public static class MaintainerInvitationFlows
    {
        // Long enough for somebody who reads mail at weekends; the code is single-use regardless.
        public static readonly TimeSpan InvitationLifetime = TimeSpan.FromDays(14);

        public const string InvitedAction = "maintainer.invited";
        public const string WithdrawnAction = "maintainer.invitation_withdrawn";
        public const string AcceptedAction = "maintainer.invitation_accepted";

        // ###########################################################################################
        // Invites an address to maintain a system. Inviting the same address to the same system
        // again replaces the open invitation - the old code stops working and a new one is mailed -
        // which is what "send it again" should do when the first mail was lost.
        // ###########################################################################################
        public static async Task<MaintainerInvitationOutcome> InviteAsync(
            ReviewAccess? actor,
            string? systemId,
            string? email,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inTree,
            IAccountStore accounts,
            ISubmissionStore submissions,
            IEmailSender mailer,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inTree);
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(mailer);

            if (!ReviewAuthority.CanAdminister(actor))
                return MaintainerInvitationOutcome.Forbidden();

            (string Manufacturer, string Hardware, string Board)? identity =
                await MaintainerAssignmentFlows.FindSystemAsync(systemId, inTree, submissions, cancellationToken);

            if (identity is null)
                return MaintainerInvitationOutcome.NotFound("No such system.");

            string address = email?.Trim() ?? string.Empty;

            if (!AccountRules.IsPlausibleEmail(address))
                return MaintainerInvitationOutcome.Refused("That does not look like an email address.");

            string normalised = AccountRules.NormaliseEmail(address);

            if (await accounts.FindByNormalisedEmailAsync(normalised, cancellationToken) is not null)
                return MaintainerInvitationOutcome.Refused(MaintainerInvitationRules.HasAccountMessage(address));

            // The invitation's foreign key needs the system's row; a shipped board nobody has
            // submitted to has none yet.
            await submissions.EnsureSystemAsync(
                systemId!, identity.Value.Manufacturer, identity.Value.Hardware, identity.Value.Board,
                SystemDescriptorRules.SystemOrigin.Shipped, now, cancellationToken);

            // Asked again: the earlier code stops working, so only the newest mail is good.
            foreach (MaintainerInvitationRecord earlier in (await accounts.ListOpenInvitationsAsync(now, cancellationToken))
                .Where(open => string.Equals(open.SystemId, systemId, StringComparison.Ordinal) && open.NormalisedEmail == normalised))
            {
                await accounts.WithdrawInvitationAsync(earlier.Id, now, cancellationToken);
            }

            string code = SecureToken.Create();

            await accounts.CreateInvitationAsync(
                new NewMaintainerInvitation(
                    systemId!, address, normalised, SecureToken.Hash(code), actor!.Account.Id,
                    now, now + MaintainerInvitationFlows.InvitationLifetime),
                cancellationToken);

            await mailer.SendAsync(
                EmailTemplates.MaintainerInvitation(
                    address,
                    MaintainerInvitationRules.SystemName(identity.Value.Manufacturer, identity.Value.Hardware, identity.Value.Board),
                    actor.Account.DisplayName,
                    code,
                    (int)MaintainerInvitationFlows.InvitationLifetime.TotalDays),
                cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(actor.Account.Id, actor.Account.Email, MaintainerInvitationFlows.InvitedAction, systemId, address, now),
                cancellationToken);

            return MaintainerInvitationOutcome.Done(
                $"An invitation is on its way to {address}. Its code works for {(int)MaintainerInvitationFlows.InvitationLifetime.TotalDays} days.");
        }

        // ###########################################################################################
        // Withdraws an invitation nobody has accepted: its code stops working.
        // ###########################################################################################
        public static async Task<MaintainerInvitationOutcome> WithdrawAsync(
            ReviewAccess? actor,
            long invitationId,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanAdminister(actor))
                return MaintainerInvitationOutcome.Forbidden();

            MaintainerInvitationRecord? invitation = await accounts.FindInvitationByIdAsync(invitationId, cancellationToken);

            if (invitation is null)
                return MaintainerInvitationOutcome.NotFound("No such invitation.");

            if (!invitation.IsOpenAt(now))
                return MaintainerInvitationOutcome.Refused("That invitation has already been accepted, withdrawn or has expired.");

            await accounts.WithdrawInvitationAsync(invitationId, now, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(actor!.Account.Id, actor.Account.Email, MaintainerInvitationFlows.WithdrawnAction,
                    invitation.SystemId, invitation.Email, now),
                cancellationToken);

            return MaintainerInvitationOutcome.Done($"The invitation to {invitation.Email} is withdrawn - its code no longer works.");
        }

        // ###########################################################################################
        // Accepts an invitation: the code from the mail, and the name and password the person picks.
        // ###########################################################################################
        public static async Task<InvitationAcceptOutcome> AcceptAsync(
            string? code,
            string? displayName,
            string? password,
            IAccountStore accounts,
            Argon2PasswordHasher hasher,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(hasher);

            if (string.IsNullOrWhiteSpace(code))
                return InvitationAcceptOutcome.Invalid("Paste the invitation code from the email.");

            // The cheap checks first, so a typo in the password costs no lookup.
            var failures = new List<string>();
            failures.AddRange(AccountRules.ValidateDisplayName(displayName));
            failures.AddRange(AccountRules.ValidatePassword(password));

            if (failures.Count > 0)
                return InvitationAcceptOutcome.Rejected(failures);

            MaintainerInvitationRecord? invitation =
                await accounts.FindInvitationByHashAsync(SecureToken.Hash(code.Trim()), cancellationToken);

            if (invitation is null)
                return InvitationAcceptOutcome.Invalid("That invitation code is not valid. Check that the whole code was pasted.");

            if (invitation.AcceptedUtc is not null)
                return InvitationAcceptOutcome.Invalid("That invitation has already been used. Sign in with your email address and password.");

            if (invitation.WithdrawnUtc is not null)
                return InvitationAcceptOutcome.Invalid("That invitation has been withdrawn. Ask the administrator for a new one.");

            if (invitation.ExpiresUtc <= now)
                return InvitationAcceptOutcome.Invalid("That invitation has expired. Ask the administrator for a new one.");

            if (await accounts.FindByNormalisedEmailAsync(invitation.NormalisedEmail, cancellationToken) is not null)
                return InvitationAcceptOutcome.Invalid(MaintainerInvitationRules.AlreadyHasAccountMessage);

            // Every system this address is invited to - accepted together, since one account
            // answers them all.
            List<string> systems = (await accounts.ListOpenInvitationsAsync(now, cancellationToken))
                .Where(open => open.NormalisedEmail == invitation.NormalisedEmail)
                .Select(open => open.SystemId)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            // Only now, holding a real open invitation: the memory-hard hash (see the header).
            long? accountId = await accounts.AcceptInvitationsAsync(
                new NewAccount(invitation.Email, invitation.NormalisedEmail, hasher.Hash(password!), displayName!.Trim(), now),
                now,
                cancellationToken);

            if (accountId is null)
                return InvitationAcceptOutcome.Invalid(MaintainerInvitationRules.AlreadyHasAccountMessage);

            // One row per system, under the system - its history says who became its maintainer.
            foreach (string system in systems)
            {
                await accounts.WriteAuditAsync(
                    new AuditEntry(accountId, invitation.Email, MaintainerInvitationFlows.AcceptedAction,
                        system, $"account {accountId} ({invitation.Email})", now),
                    cancellationToken);
            }

            return InvitationAcceptOutcome.Accepted(new AcceptInvitationAnswer(
                invitation.Email,
                systems,
                MaintainerInvitationRules.AcceptedMessage(systems)));
        }
    }

    // ###########################################################################################
    // The sentences, pure, so they are tested - one of them is read by a stranger who has just
    // been invited and knows nothing of how this works.
    // ###########################################################################################
    public static class MaintainerInvitationRules
    {
        public const string AlreadyHasAccountMessage =
            "An account with this address already exists. Sign in with it, and ask the administrator to add you to the system.";

        public static string HasAccountMessage(string email) =>
            $"{email} already has an account - choose it from the list of accounts instead of inviting it.";

        // "Commodore / C64 / 250407", as the Systems screen names it.
        public static string SystemName(string manufacturer, string hardware, string board) =>
            string.Join(" / ", new[] { manufacturer, hardware, board }.Where(part => !string.IsNullOrWhiteSpace(part)));

        public static string AcceptedMessage(IReadOnlyList<string> systemIds)
        {
            ArgumentNullException.ThrowIfNull(systemIds);

            string what = systemIds.Count switch
            {
                0 => "Your account is ready.",
                1 => $"Your account is ready, and you now maintain {systemIds[0]}.",
                _ => $"Your account is ready, and you now maintain {string.Join(", ", systemIds.Take(systemIds.Count - 1))} and {systemIds[^1]}."
            };

            return $"{what} Sign in with your email address and the password you chose.";
        }
    }

    public sealed record MaintainerInvitationOutcome(
        bool IsDone,
        string Message,
        bool IsForbidden = false,
        bool IsNotFound = false)
    {
        public static MaintainerInvitationOutcome Done(string message) => new(true, message);

        public static MaintainerInvitationOutcome Forbidden() =>
            new(false, "Only an administrator can invite maintainers.", IsForbidden: true);

        public static MaintainerInvitationOutcome NotFound(string error) => new(false, error, IsNotFound: true);

        public static MaintainerInvitationOutcome Refused(string error) => new(false, error);
    }

    public sealed record InvitationAcceptOutcome(
        AcceptInvitationAnswer? Answer,
        string Message,
        IReadOnlyList<string> Errors)
    {
        public bool IsAccepted => this.Answer is not null;

        public static InvitationAcceptOutcome Accepted(AcceptInvitationAnswer answer) => new(answer, answer.Message, []);

        public static InvitationAcceptOutcome Invalid(string message) => new(null, message, []);

        // The name or the password broke a rule; every broken rule is listed.
        public static InvitationAcceptOutcome Rejected(IReadOnlyList<string> errors) =>
            new(null, string.Join(" ", errors), errors);
    }
}
