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
        public const int MinimumPasswordLength = 12;

        // Nobody has a legitimate 200-character password, and unbounded input reaches the hasher,
        // which allocates per call. The limit is generous enough for any passphrase or password
        // manager output and small enough to be harmless.
        public const int MaximumPasswordLength = 256;

        // RFC 5321's limit on a full address. The database column matches.
        // Kept as an alias so existing references read naturally; the value lives in CRT.Data.
        public const int MaximumEmailLength = EmailAddressRules.MaximumLength;

        public const int MinimumDisplayNameLength = 2;
        public const int MaximumDisplayNameLength = 100;

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
        // Checks a display name. This is the name shown beside a contribution, so it has to be
        // something a human can read, but there is no reason to be restrictive about scripts -
        // contributors are worldwide.
        // ###########################################################################################
        public static IReadOnlyList<string> ValidateDisplayName(string? displayName)
        {
            var failures = new List<string>();

            if (string.IsNullOrWhiteSpace(displayName))
            {
                failures.Add("A display name is required. It is shown beside your contributions.");
                return failures;
            }

            string trimmed = displayName.Trim();

            if (trimmed.Length < AccountRules.MinimumDisplayNameLength)
            {
                failures.Add(string.Create(
                    CultureInfo.InvariantCulture,
                    $"The display name must be at least {AccountRules.MinimumDisplayNameLength} characters."));
            }

            if (trimmed.Length > AccountRules.MaximumDisplayNameLength)
            {
                failures.Add(string.Create(
                    CultureInfo.InvariantCulture,
                    $"The display name must be at most {AccountRules.MaximumDisplayNameLength} characters."));
            }

            // Control characters would let a name break a log line or a mail header. Ordinary
            // letters, marks and punctuation from any script are fine.
            foreach (char character in trimmed)
            {
                if (char.IsControl(character))
                {
                    failures.Add("The display name cannot contain control characters.");
                    break;
                }
            }

            return failures;
        }
    }
}
