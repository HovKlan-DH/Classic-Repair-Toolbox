using System.Globalization;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // The rules about what an account may be: which addresses are acceptable, how they are
    // normalised, and what counts as an acceptable password. Pure, so every rule is a unit test.
    //
    // KNOW WHO IS READING. CRT's users are hobbyists repairing their own vintage computers, not
    // employees under a corporate password policy. Every message here is written to be read by
    // someone at a bench, and the password rule is deliberately LENGTH-led rather than a
    // composition rule demanding symbols and digits - NIST SP 800-63B withdrew those years ago
    // because they produce "Password1!" and a sticky note, not entropy.
    // ###########################################################################################
    public static class AccountRules
    {
        // Twelve, not eight. A long passphrase is both stronger and easier to remember than a
        // short mangled word, and this audience logs in rarely enough that typing a few more
        // characters costs nothing.
        // CRT.Data's number since 2026-11-25, so the sentence CRT's account window and the account web pages say beside the
        // box (AccountWording.PasswordLine) is held to it.
        public const int MinimumPasswordLength = AccountWording.MinimumPasswordLength;

        // Nobody has a legitimate 200-character password, and unbounded input reaches the hasher,
        // which allocates per call. The limit is generous enough for any passphrase or password
        // manager output and small enough to be harmless.
        public const int MaximumPasswordLength = 256;

        // RFC 5321's limit on a full address. The database column matches.
        // Kept as an alias so existing references read naturally; the value lives in CRT.Data.
        public const int MaximumEmailLength = EmailAddressRules.MaximumLength;


        // ###########################################################################################
        // Email rules live in CRT.Data (EmailAddressRules) because BOTH SIDES need them: the app
        // validates the address typed into the Submit dialog before enabling the button, and the
        // server validates contributors' contact addresses and accounts' addresses. Two copies
        // would drift, and the drift would show as the app accepting something the server refuses
        // - which reads to a user as the server being broken.
        //
        // These stay here as the names the account code already uses, so a reader looking for
        // "the account rules" finds them, and there is still exactly one implementation.
        // ###########################################################################################
        public static string NormaliseEmail(string email) => EmailAddressRules.Normalise(email);

        public static bool IsPlausibleEmail(string? email) => EmailAddressRules.IsPlausible(email);

        // ###########################################################################################
        // Checks a password against the policy, returning every problem at once.
        //
        // Every problem at once, for the same reason ServerOptionsValidator accumulates failures:
        // being told one rule, fixing it, and being told the next is a miserable way to choose a
        // password, and it pushes people towards the shortest thing that passes.
        //
        // WHAT IS DELIBERATELY NOT HERE: no required character classes, no forced expiry, no
        // prohibition on repeated characters. All three are withdrawn NIST guidance, and all three
        // reliably produce weaker passwords by pushing people into predictable patterns.
        // ###########################################################################################
        public static IReadOnlyList<string> ValidatePassword(string? password)
        {
            var failures = new List<string>();

            if (string.IsNullOrEmpty(password))
            {
                failures.Add("A password is required.");
                return failures;
            }

            if (password.Length < AccountRules.MinimumPasswordLength)
            {
                failures.Add(string.Create(
                    CultureInfo.InvariantCulture,
                    $"The password must be at least {AccountRules.MinimumPasswordLength} characters. " +
                    $"A few ordinary words together make a strong password that is easy to remember."));
            }

            if (password.Length > AccountRules.MaximumPasswordLength)
            {
                failures.Add(string.Create(
                    CultureInfo.InvariantCulture,
                    $"The password must be at most {AccountRules.MaximumPasswordLength} characters."));
            }

            // Whitespace-only passwords pass a naive length check and are almost certainly a
            // paste accident.
            if (password.Trim().Length == 0)
                failures.Add("The password cannot be only spaces.");

            return failures;
        }

        // ###########################################################################################
        // Checks a display name being SET - at registration, under "My account" and accepting an invitation. CRT.Data's
        // DisplayNames is the rule (owner requests, 2026-11-25: unique, "only normal A-Z + 0-9 and then some common special
        // charatcers like: +-/@.[]()", "max length 30"), so CRT's account window says the very sentence as the name is typed.
        // Any script was taken until then, 2 to 100 characters with no control character. Whether another account has the
        // name is asked apart (IAccountStore.IsDisplayNameTakenAsync) - this needs no database. A control character, which
        // would let a name break a log line or a mail header, is outside the rule's characters too.
        // ###########################################################################################
        public static IReadOnlyList<string> ValidateDisplayName(string? displayName) =>
            DisplayNames.Problem(displayName) is string problem ? [problem] : [];
    }
}
