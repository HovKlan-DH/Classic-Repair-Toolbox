using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW A COLLECTION'S LINK IS WRITTEN FOR WHERE IT IS PASTED (owner decision, 2026-11-28: "OK to your proposal for the forum
    // button" - forums do not differ by site, only by how they write a link, so a format each, never a button per forum):
    // BBCode for Lemon64, Forum64 and most forums, Markdown for Reddit, Discord and GitHub, and the plain link after a line of
    // text for everywhere else. ONE rule for CRT's "Copy link" (CollectionSharing.ForumText) and the page's "Link" panel
    // (CollectionPages) - "Copy for Lemon64", "Copy for Forum64" and "Copy for Reddit" until then. Stored by its name (the
    // settings' `collectionLinkFormat`), so never rename a member.
    // ###########################################################################################
    public enum CollectionLinkFormat
    {
        BbCode,
        Markdown,
        Plain
    }

    public static class CollectionLinkFormats
    {
        public static IReadOnlyList<CollectionLinkFormat> All { get; } =
            [CollectionLinkFormat.BbCode, CollectionLinkFormat.Markdown, CollectionLinkFormat.Plain];

        // The link to `url` with `text` as its words, as `format` writes it. A square bracket in the words is written round:
        // in BBCode a "[/url]" typed in a title would end the link, and in Markdown a "]" ends its words.
        public static string Write(CollectionLinkFormat format, string url, string text)
        {
            ArgumentNullException.ThrowIfNull(url);

            string words = (text ?? string.Empty).Replace("[", "(", StringComparison.Ordinal).Replace("]", ")", StringComparison.Ordinal);

            return format switch
            {
                CollectionLinkFormat.Markdown => $"[{words}]({url})",
                CollectionLinkFormat.Plain => words.Length == 0 ? url : $"{words}: {url}",
                _ => $"[url={url}]{words}[/url]"
            };
        }

        // Its name, as a person picks it.
        public static string NameOf(CollectionLinkFormat format) => format switch
        {
            CollectionLinkFormat.Markdown => "Markdown",
            CollectionLinkFormat.Plain => "Plain link",
            _ => "BBCode"
        };

        // Where it works - the examples beside its name.
        public static string WhereOf(CollectionLinkFormat format) => format switch
        {
            CollectionLinkFormat.Markdown => "Reddit, Discord, GitHub and Discourse",
            CollectionLinkFormat.Plain => "Facebook groups, email and anything else",
            _ => "Lemon64, Forum64, AtariAge, VOGONS and most other forums"
        };

        // A stored name read back - anything else, or none, is BBCode, the most forums' format.
        public static CollectionLinkFormat Parse(string? stored) =>
            Enum.TryParse(stored, ignoreCase: true, out CollectionLinkFormat format) && Enum.IsDefined(format) && !int.TryParse(stored, out _)
                ? format
                : CollectionLinkFormat.BbCode;
    }
}
