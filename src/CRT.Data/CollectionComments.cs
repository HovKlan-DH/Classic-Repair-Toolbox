using System;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // COMMENTS ON A SHARED COLLECTION'S CAPTURES (step 3, 2026-11-26; owner decisions, 2026-11-25: "registered users comment
    // on each image", and 2026-11-26: the owner mailed for every comment, those who comment not; the writer edits and
    // deletes their own; the owner hides any on the collection). Every rule both ends keep:
    //
    //   - a comment is on ONE capture, by an account with a confirmed address - signed in on the page, or in CRT;
    //   - plain text, 1 to MaximumLength characters - shown as typed, never read as markup or a link;
    //   - its writer's name beside it, "Owner" by the collection's owner's, "Administrator" by an administrator's;
    //   - hidden by the owner: gone from the page, still listed for the owner; deleted by its writer: gone.
    // ###########################################################################################
    public static class CollectionComments
    {
        public const int MaximumLength = 2000;

        // How many comments an account writes in a day, everywhere together - far above a person helping, a wall to a script.
        public const int MaximumPerDay = 100;

        public const string OwnerBadge = "Owner";

        public const string AdministratorBadge = "Administrator";

        public const string Missing = "Write the comment first.";

        // What is wrong with a comment's text - null when nothing is. Measured AS KEPT (Tidied): a browser posts each line end as
        // "\r\n" but its maxlength counts it as one, so a comment the page let through was refused as too long (code review,
        // 2026-11-27).
        public static string? Problem(string? body)
        {
            string text = body is null ? string.Empty : CollectionComments.Tidied(body);

            if (text.Length == 0)
                return CollectionComments.Missing;

            return text.Length > CollectionComments.MaximumLength
                ? $"A comment is at most {CollectionComments.MaximumLength} characters - this one has {text.Length}."
                : null;
        }

        // A comment's text as kept: trimmed, its line ends "\n".
        public static string Tidied(string body) => body.Trim().Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');

        // When it was written, as a person reads it beside the comment: "2026-November-26 14:05 UTC".
        public static string WhenText(DateTimeOffset when) =>
            when.ToUniversalTime().ToString("yyyy-MMMM-d HH:mm", CultureInfo.InvariantCulture) + " UTC";
    }
}
