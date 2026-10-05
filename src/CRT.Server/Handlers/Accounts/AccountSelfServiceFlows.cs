using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // A SIGNED-IN MAINTAINER CHANGING THEIR OWN ACCOUNT (owner request, 2026-10-03: the maintainer
    // "can edit his/her own email address and name"). The Maintainer tab's "Your account" window.
    //
    // Every method takes the request's bearer token and authenticates it itself, so nothing here
    // can be reached without a live session - and the session is what an EMAIL or PASSWORD change
    // keeps while signing out every other one.
    //
    // *** THE SESSION IS ENOUGH - NO CURRENT PASSWORD (owner decision, 2026-10-03: "no need for
    // that, as I see it as you are already logged in"). *** The first version asked for it before
    // an address or password change, so that a session token copied off disk could not take the
    // account over. The project owner chose against it, knowing that; it is an ACCEPTED RISK,
    // recorded in NewContributeStrategy.md's Threat 2. What still stands between a copied token
    // and a lost account:
    //
    //   EMAIL - a code mailed to the NEW address must come back before anything changes, so the
    //   new mailbox is proved and a typo cannot strand the account at an address nobody reads; and
    //   the OLD address is told once it has changed, naming whom to write to. Only the capitals
    //   changing needs no code: it is the same mailbox (AccountRules.NormaliseEmail).
    //
    //   PASSWORD - the address is unchanged, so "I forgot my password" still reaches the owner, and
    //   a reset signs out every session, the copier's included. The account's address is told.
    //
    // *** BOTH SPEND THE PER-ADDRESS MAIL BUDGET *** (AccountFlows.CheckMailBudgetAsync, the one
    // registration and "I forgot my password" spend). Each sends a mail; and a password change
    // also runs the ~128 MiB Argon2 hasher, which with no password check in front of it would
    // otherwise be callable without limit - the limiter-before-hasher rule of AccountFlows' header.
    //
    // *** AN ADDRESS THAT ALREADY HAS AN ACCOUNT GETS THE SAME ANSWER AS A FREE ONE (owner
    // decision, 2026-10-03). *** "A code is on its way" either way; what differs is the MAIL - the
    // free address gets the code, the taken one is told somebody tried to move an account to it.
    // The same property registration keeps (EmailTemplates' header): only the mailbox owner learns
    // whether the address is registered.
    //
    // *** AN EMAIL OR PASSWORD CHANGE SIGNS OUT EVERY OTHER SESSION (owner decision, 2026-10-03:
    // "Once changed, then it should logoff from anywhere the old session/account is used"), and
    // keeps this one. *** If the account had been taken over, that is what ends it; and CRT on
    // another computer then asks for the new details instead of showing the old ones.
    //
    // Pure, like every flow: the store and the mailer are arguments, the endpoints a rim.
    // ###########################################################################################
    public static class AccountSelfServiceFlows
    {
        // Long, like the verification link: the code is useless without the maintainer's own
        // session (the confirmation is signed in), and somebody may ask in the evening and read
        // the mail the next day.
        public static readonly TimeSpan EmailChangeCodeLifetime = TimeSpan.FromHours(24);

        public const string NameChangedAction = "account.name_changed";
        public const string EmailChangeRequestedAction = "account.email_change_requested";
        public const string EmailChangeTakenAction = "account.email_change_taken";
        public const string EmailChangedAction = "account.email_changed";
        public const string PasswordChangedAction = "password.changed";

        // The sentences a maintainer reads, kept here so the tests assert the same words.
        public const string NotAnEmailAddress = "That does not look like an email address.";
        public const string AlreadyYourAddress = "That is already your email address.";
        public const string PasteTheCode = "Type in the code from the email.";
        public const string CodeNotValid = "That code is not valid. Check that the whole code was typed in.";
        public const string CodeAlreadyUsed = "That code has already been used. Ask for a new one if your address has not changed.";
        public const string CodeExpired = "That code has expired. Ask for a new one.";
        public const string AddressTakenMeanwhile = "Another account has taken that address since the code was sent, so your address was not changed.";

        public const string PasswordChangedMessage =
            "Your password is changed. CRT on any other computer will ask you to sign in again.";

        // The answer to every new-address request that went through, taken address or not.
        public static string CodeSentMessage(string address) =>
            $"A code is on its way to {address}. Your address stays as it is until the code is typed in.";

        public static string NameChangedMessage(string name) => $"Your name is now {name}.";

        public static string EmailChangedMessage(string address) =>
            $"Your email address is now {address}. CRT on any other computer will ask you to sign in again, with this address.";

        public static string EmailRespelledMessage(string address) => $"Your email address is now written {address}.";

        // ###########################################################################################
        // The account as GET /api/accounts/me answers it - and as every change below hands it back,
        // so CRT shows what the server now holds rather than what it sent.
        // ###########################################################################################
        public static async Task<AccountAnswer> DescribeAsync(
            AccountRecord account,
            IAccountStore store,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(account);
            ArgumentNullException.ThrowIfNull(store);

            // The systems whose pools this account is in. For an administrator, who reviews every
            // system anyway, only those they were named a maintainer of (2026-10-05) - often none.
            IReadOnlySet<string> maintainerOf = await store.GetReviewedSystemIdsAsync(account.Id, cancellationToken);

            return new AccountAnswer(
                account.Id,
                account.Email,
                account.DisplayName,
                account.IsVerified,
                account.IsAdministrator,
                maintainerOf.OrderBy(id => id, StringComparer.Ordinal).ToList(),
                account.CreatedUtc,
                account.LastLoginUtc);
        }

        // ###########################################################################################
        // NAME.
        // ###########################################################################################
        public static async Task<AccountChangeOutcome> ChangeNameAsync(
            string? bearerToken,
            string? displayName,
            IAccountStore store,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(options);

            SignedIn? signedIn = await AccountSelfServiceFlows.SignedInAsync(bearerToken, store, options, now, cancellationToken);

            if (signedIn is null)
                return AccountChangeOutcome.NotSignedIn();

            AccountRecord account = signedIn.Account;

            IReadOnlyList<string> failures = AccountRules.ValidateDisplayName(displayName);

            if (failures.Count > 0)
                return AccountChangeOutcome.Rejected(failures);

            string name = displayName!.Trim();

            // Nothing to write, and nothing worth an audit row.
            if (string.Equals(name, account.DisplayName, StringComparison.Ordinal))
            {
                return AccountChangeOutcome.Done(new AccountChangeAnswer(
                    AccountSelfServiceFlows.NameChangedMessage(name),
                    await AccountSelfServiceFlows.DescribeAsync(account, store, cancellationToken)));
            }

            await store.SetDisplayNameAsync(account.Id, name, cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, account.Email, AccountSelfServiceFlows.NameChangedAction,
                    $"account:{account.Id}", $"\"{account.DisplayName}\" -> \"{name}\"", now),
                cancellationToken);

            return AccountChangeOutcome.Done(new AccountChangeAnswer(
                AccountSelfServiceFlows.NameChangedMessage(name),
                await AccountSelfServiceFlows.DescribeCurrentAsync(account with { DisplayName = name }, store, cancellationToken)));
        }

        // ###########################################################################################
        // EMAIL, step 1: the new address. Mails the code - or, for an address another account
        // holds, the "somebody tried" notice - and answers the same either way (see the header).
        // Only the capitals changing is done at once.
        // ###########################################################################################
        public static async Task<AccountChangeOutcome> RequestEmailChangeAsync(
            string? bearerToken,
            string? newEmail,
            string? ipAddress,
            IAccountStore store,
            IEmailSender mailer,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);
            ArgumentNullException.ThrowIfNull(options);

            SignedIn? signedIn = await AccountSelfServiceFlows.SignedInAsync(bearerToken, store, options, now, cancellationToken);

            if (signedIn is null)
                return AccountChangeOutcome.NotSignedIn();

            AccountRecord account = signedIn.Account;
            string address = newEmail?.Trim() ?? string.Empty;

            if (!AccountRules.IsPlausibleEmail(address))
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.NotAnEmailAddress);

            string normalised = AccountRules.NormaliseEmail(address);
            bool sameMailbox = string.Equals(normalised, account.NormalisedEmail, StringComparison.Ordinal);

            if (sameMailbox && string.Equals(address, account.Email, StringComparison.Ordinal))
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.AlreadyYourAddress);

            // ---- Only the capitals differ: the same mailbox, nothing to prove. ----
            if (sameMailbox)
            {
                if (!await store.SetEmailAsync(account.Id, address, normalised, cancellationToken))
                    return AccountChangeOutcome.Refused(AccountSelfServiceFlows.AddressTakenMeanwhile);

                await store.WriteAuditAsync(
                    new AuditEntry(account.Id, address, AccountSelfServiceFlows.EmailChangedAction,
                        $"account:{account.Id}", $"{account.Email} -> {address}", now),
                    cancellationToken);

                return AccountChangeOutcome.Done(new AccountChangeAnswer(
                    AccountSelfServiceFlows.EmailRespelledMessage(address),
                    await AccountSelfServiceFlows.DescribeCurrentAsync(account with { Email = address }, store, cancellationToken)));
            }

            // ---- A mail goes out from here on, to an address the caller names. ----
            //
            // The registration and reset budget: signed in or not, this is a way to put mail into
            // a mailbox somebody else owns. Answered as 429 rather than silently, since the
            // maintainer would otherwise wait for a code that is never sent.
            RateLimitVerdict mailVerdict = await AccountFlows.CheckMailBudgetAsync(ipAddress, store, now, cancellationToken);

            if (!mailVerdict.IsAllowed)
                return AccountChangeOutcome.RateLimited(mailVerdict.RetryAfter);

            // Only the newest request counts: a code mailed for an earlier address stops working,
            // whichever way this one goes.
            await store.ConsumeOutstandingTokensAsync(account.Id, TokenPurpose.EmailChange, now, cancellationToken);

            AccountRecord? holder = await store.FindByNormalisedEmailAsync(normalised, cancellationToken);

            if (holder is not null)
            {
                // The holder's own name and address - nothing the requester typed but the address
                // itself goes into a mail to somebody else.
                await mailer.SendAsync(EmailTemplates.EmailChangeAddressTaken(holder.Email, holder.DisplayName), cancellationToken);

                await store.WriteAuditAsync(
                    new AuditEntry(account.Id, account.Email, AccountSelfServiceFlows.EmailChangeTakenAction,
                        $"account:{account.Id}", $"{address} belongs to account {holder.Id}", now),
                    cancellationToken);

                return AccountChangeOutcome.Done(new AccountChangeAnswer(
                    AccountSelfServiceFlows.CodeSentMessage(address), Account: null, CodeSent: true));
            }

            string code = SecureToken.Create();

            await store.CreateTokenAsync(
                new NewAccountToken(
                    account.Id,
                    TokenPurpose.EmailChange,
                    SecureToken.Hash(code),
                    now,
                    now + AccountSelfServiceFlows.EmailChangeCodeLifetime,
                    address,
                    normalised),
                cancellationToken);

            await mailer.SendAsync(
                EmailTemplates.EmailChangeCode(
                    address,
                    account.DisplayName,
                    code,
                    (int)AccountSelfServiceFlows.EmailChangeCodeLifetime.TotalHours),
                cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, account.Email, AccountSelfServiceFlows.EmailChangeRequestedAction,
                    $"account:{account.Id}", address, now),
                cancellationToken);

            return AccountChangeOutcome.Done(new AccountChangeAnswer(
                AccountSelfServiceFlows.CodeSentMessage(address), Account: null, CodeSent: true));
        }

        // ###########################################################################################
        // EMAIL, step 2: the code from the new mailbox. Signed in as the SAME account the code was
        // made for - a code is useless in anybody else's session, and says so no differently from
        // a mistyped one.
        // ###########################################################################################
        public static async Task<AccountChangeOutcome> ConfirmEmailChangeAsync(
            string? bearerToken,
            string? code,
            IAccountStore store,
            IEmailSender mailer,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);
            ArgumentNullException.ThrowIfNull(options);

            SignedIn? signedIn = await AccountSelfServiceFlows.SignedInAsync(bearerToken, store, options, now, cancellationToken);

            if (signedIn is null)
                return AccountChangeOutcome.NotSignedIn();

            AccountRecord account = signedIn.Account;

            if (string.IsNullOrWhiteSpace(code))
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.PasteTheCode);

            AccountTokenRecord? record = await store.FindTokenByHashAsync(SecureToken.Hash(code.Trim()), cancellationToken);

            if (record is null ||
                record.Purpose != TokenPurpose.EmailChange ||
                record.AccountId != account.Id ||
                string.IsNullOrWhiteSpace(record.PendingEmail) ||
                string.IsNullOrWhiteSpace(record.PendingEmailNormalised))
            {
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.CodeNotValid);
            }

            if (record.ConsumedUtc is not null)
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.CodeAlreadyUsed);

            if (record.ExpiresUtc <= now)
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.CodeExpired);

            string oldAddress = account.Email;
            string newAddress = record.PendingEmail;

            // The unique index decides a race; this is the early, readable answer for an address
            // registered or moved to between the request and now.
            AccountRecord? holder = await store.FindByNormalisedEmailAsync(record.PendingEmailNormalised, cancellationToken);
            bool takenMeanwhile = holder is not null && holder.Id != account.Id;

            if (takenMeanwhile || !await store.SetEmailAsync(account.Id, newAddress, record.PendingEmailNormalised, cancellationToken))
            {
                await store.ConsumeTokenAsync(record.Id, now, cancellationToken);
                return AccountChangeOutcome.Refused(AccountSelfServiceFlows.AddressTakenMeanwhile);
            }

            await store.ConsumeTokenAsync(record.Id, now, cancellationToken);
            await store.ConsumeOutstandingTokensAsync(account.Id, TokenPurpose.EmailChange, now, cancellationToken);

            // A reset code mailed to the OLD address stops working too. If that mailbox is the
            // reason for the change - lost, or read by somebody else - a reset code still in it
            // would be a way back in.
            await store.ConsumeOutstandingTokensAsync(account.Id, TokenPurpose.PasswordReset, now, cancellationToken);

            // The code reaching the new mailbox proves what the verification link would have.
            if (!account.IsVerified)
                await store.SetVerifiedAsync(account.Id, cancellationToken);

            await store.RevokeOtherSessionsAsync(account.Id, signedIn.SessionId, "email changed", now, cancellationToken);

            // To the OLD address: the one notice that reaches the owner if this was not them.
            await mailer.SendAsync(EmailTemplates.EmailChanged(oldAddress, account.DisplayName, newAddress), cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, newAddress, AccountSelfServiceFlows.EmailChangedAction,
                    $"account:{account.Id}", $"{oldAddress} -> {newAddress}", now),
                cancellationToken);

            return AccountChangeOutcome.Done(new AccountChangeAnswer(
                AccountSelfServiceFlows.EmailChangedMessage(newAddress),
                await AccountSelfServiceFlows.DescribeCurrentAsync(
                    account with { Email = newAddress, NormalisedEmail = record.PendingEmailNormalised, IsVerified = true },
                    store,
                    cancellationToken)));
        }

        // ###########################################################################################
        // PASSWORD: the new one, held to the same rules as at registration. The rules are checked
        // and the mail budget spent BEFORE the hasher runs - see the header.
        // ###########################################################################################
        public static async Task<AccountChangeOutcome> ChangePasswordAsync(
            string? bearerToken,
            string? newPassword,
            string? ipAddress,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);
            ArgumentNullException.ThrowIfNull(hasher);
            ArgumentNullException.ThrowIfNull(options);

            SignedIn? signedIn = await AccountSelfServiceFlows.SignedInAsync(bearerToken, store, options, now, cancellationToken);

            if (signedIn is null)
                return AccountChangeOutcome.NotSignedIn();

            AccountRecord account = signedIn.Account;

            IReadOnlyList<string> failures = AccountRules.ValidatePassword(newPassword);

            if (failures.Count > 0)
                return AccountChangeOutcome.Rejected(failures);

            RateLimitVerdict mailVerdict = await AccountFlows.CheckMailBudgetAsync(ipAddress, store, now, cancellationToken);

            if (!mailVerdict.IsAllowed)
                return AccountChangeOutcome.RateLimited(mailVerdict.RetryAfter);

            await store.SetPasswordHashAsync(account.Id, hasher.Hash(newPassword!), cancellationToken);

            // A reset code already in a mailbox must not set the password back over this one.
            await store.ConsumeOutstandingTokensAsync(account.Id, TokenPurpose.PasswordReset, now, cancellationToken);

            await store.RevokeOtherSessionsAsync(account.Id, signedIn.SessionId, "password changed", now, cancellationToken);

            await mailer.SendAsync(EmailTemplates.PasswordChanged(account.Email, account.DisplayName), cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, account.Email, AccountSelfServiceFlows.PasswordChangedAction,
                    $"account:{account.Id}", null, now),
                cancellationToken);

            return AccountChangeOutcome.Done(new AccountChangeAnswer(
                AccountSelfServiceFlows.PasswordChangedMessage,
                await AccountSelfServiceFlows.DescribeCurrentAsync(account, store, cancellationToken)));
        }

        // -------------------------------------------------------------------------------------
        // Helpers.
        // -------------------------------------------------------------------------------------

        private sealed record SignedIn(AccountRecord Account, long SessionId);

        // The account AND the session the token names - the session is the one an email or
        // password change keeps. AuthenticateAsync makes every check (revoked, rotated, expired,
        // locked) and slides the expiry; the session is then read for its id alone.
        private static async Task<SignedIn?> SignedInAsync(
            string? bearerToken,
            IAccountStore store,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            AccountRecord? account = await AccountFlows.AuthenticateAsync(bearerToken, store, now, cancellationToken, options);

            if (account is null)
                return null;

            SessionRecord? session = await store.FindSessionByHashAsync(SecureToken.Hash(bearerToken!), cancellationToken);

            return session is null ? null : new SignedIn(account, session.Id);
        }

        // The account as the store now holds it, after a change; `expected` if it cannot be read
        // back (it was just written, so that is a database gone away mid-request).
        private static async Task<AccountAnswer> DescribeCurrentAsync(
            AccountRecord expected,
            IAccountStore store,
            CancellationToken cancellationToken)
        {
            AccountRecord current = await store.FindByIdAsync(expected.Id, cancellationToken) ?? expected;

            return await AccountSelfServiceFlows.DescribeAsync(current, store, cancellationToken);
        }
    }

    // ###########################################################################################
    // What a change came to. Exactly one of: Done (with the answer), NotSignedIn (401), RateLimited
    // (429 with RetryAfter), or refused (400) - with Errors when a name or password broke rules,
    // else Message.
    // ###########################################################################################
    public sealed record AccountChangeOutcome(
        AccountChangeAnswer? Answer,
        string Message,
        IReadOnlyList<string> Errors,
        bool IsNotSignedIn = false,
        TimeSpan? RetryAfter = null)
    {
        public bool IsDone => this.Answer is not null;

        public static AccountChangeOutcome Done(AccountChangeAnswer answer) => new(answer, answer.Message, []);

        public static AccountChangeOutcome NotSignedIn() => new(null, string.Empty, [], IsNotSignedIn: true);

        public static AccountChangeOutcome RateLimited(TimeSpan retryAfter) =>
            new(null, "Too many attempts. Wait a while and try again.", [], RetryAfter: retryAfter);

        public static AccountChangeOutcome Refused(string message) => new(null, message, []);

        public static AccountChangeOutcome Rejected(IReadOnlyList<string> errors) =>
            new(null, string.Join(" ", errors), errors);
    }
}
