using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Classifies a contributor's free-text contact string and turns it into an openable href.
    //
    // Extracted from TabAbout. Contact strings come from contributed credit data, so they are
    // whatever a contributor typed: a bare "www." host, a full URL, an email address, or plain
    // text that is none of those and must be left alone rather than guessed at.
    // ###########################################################################################
    public static class ContactLinkFormatter
    {
        // ###########################################################################################
        // Returns true when the contact string looks like a web URL.
        // ###########################################################################################
        public static bool IsContactWebUrl(string contact)
            => contact.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || contact.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            || contact.StartsWith("www.", StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Returns true when the contact string looks like an email address. The no-space rule is
        // what stops a sentence that happens to mention an address being treated as one.
        // ###########################################################################################
        public static bool IsContactEmail(string contact)
            => contact.Contains('@') && !contact.Contains(' ');

        // ###########################################################################################
        // Builds the href to open from contact text. Email wins over the URL check, so an address
        // is never turned into an https link; a bare "www." host is promoted to https.
        // ###########################################################################################
        public static string BuildContactHref(string contact)
        {
            if (IsContactEmail(contact))
                return $"mailto:{contact}";
            if (contact.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
                return $"https://{contact}";
            return contact;
        }
    }
}
