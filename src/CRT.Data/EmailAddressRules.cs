using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What counts as a usable email address, and how one is normalised for comparison.
    //
    // IN CRT.DATA BECAUSE BOTH SIDES NEED IT. The server validates a contributor's contact address
    // and an account's address; the app validates the address typed into the Submit dialog before
    // enabling the button. Two copies of this rule would drift, and the drift would show up as the
    // app cheerfully accepting something the server then refuses - which reads to the user as the
    // server being broken.
    //
    // This is exactly the shared-rule case CRT.Data was extracted for in Phase 1.
    // ###########################################################################################
    public static class EmailAddressRules
    {
        // RFC 5321's limit on a full address. The database columns match.
        public const int MaximumLength = 320;

        // ###########################################################################################
        // Normalises an address for COMPARISON and for a unique index. Trim, then lowercase.
        //
        // WHY LOWERCASE THE WHOLE THING, local part included. Strictly, RFC 5321 makes the local
        // part case-SENSITIVE, so "Dennis@example.com" and "dennis@example.com" are formally two
        // different mailboxes. In practice no mail system CRT's users actually have treats them as
        // different, and honouring the letter of the standard would let one person register twice
        // by changing a capital.
        //
        // INVARIANT CULTURE, not the current culture. Turkish lowercases 'I' to a dotless 'i', so
        // a culture-sensitive lowercase would normalise the same address differently depending on
        // the machine's locale - and on the server this value is a database key.
        // ###########################################################################################
        public static string Normalise(string email)
        {
            ArgumentNullException.ThrowIfNull(email);

            return email.Trim().ToLowerInvariant();
        }

        // ###########################################################################################
        // Is this a usable email address?
        //
        // DELIBERATELY NOT A FULL RFC 5322 VALIDATOR. That grammar admits comments, quoted strings
        // and nested folding whitespace, and every regex claiming to implement it is either wrong
        // or unreadable. The only thing that actually proves an address works is sending mail to
        // it and having someone read it. So this checks only the shape that catches honest typos.
        //
        // What it rejects is therefore narrow and defensible: no '@', more than one '@', nothing
        // before or after it, no dot in the domain, whitespace anywhere, or over-length.
        // ###########################################################################################
        public static bool IsPlausible(string? email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            string trimmed = email.Trim();

            if (trimmed.Length > EmailAddressRules.MaximumLength)
                return false;

            // Whitespace inside an address is legal only inside a quoted local part, which nobody
            // in this audience has. Rejecting it catches a pasted "Name <a@b.com>" outright.
            foreach (char character in trimmed)
            {
                if (char.IsWhiteSpace(character))
                    return false;
            }

            string[] parts = trimmed.Split('@');

            if (parts.Length != 2)
                return false;

            string local = parts[0];
            string domain = parts[1];

            if (local.Length == 0 || domain.Length == 0)
                return false;

            // A domain with no dot is either a local hostname or a typo. Mail to it will not leave
            // the machine.
            int lastDot = domain.LastIndexOf('.');

            if (lastDot <= 0 || lastDot == domain.Length - 1)
                return false;

            // Consecutive dots are invalid in a domain, and a leading dot is too.
            if (domain.StartsWith('.') || domain.Contains(".."))
                return false;

            return true;
        }
    }
}
