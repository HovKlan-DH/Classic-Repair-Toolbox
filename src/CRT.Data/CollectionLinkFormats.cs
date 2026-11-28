using System;
using System.Collections.Generic;
using System.Net;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW A COLLECTION'S LINK IS WRITTEN FOR WHERE IT IS PASTED (owner decision, 2026-11-28: "OK to your proposal for the forum
    // button" - forums do not differ by site, only by how they write a link, so a format each, never a button per forum):
    // BBCode for Lemon64, Forum64 and most forums, Markdown for Reddit, Discord and GitHub, HTML for a web page (owner request,
    // the same day: "add an option of "HTML" for the link options and embed the link in "<a href>""), and the plain link after a
    // line of text for everywhere else. ONE rule for CRT's "Copy link" (CollectionSharing.ForumText) and the page's "Link" panel
    // (CollectionPages) - "Copy for Lemon64", "Copy for Forum64" and "Copy for Reddit" until then. Stored by its name (the
    // settings' `collectionLinkFormat`), so never rename a member.
    // ###########################################################################################
    public enum CollectionLinkFormat
    {
        BbCode,
        Markdown,
        Plain,
        Html
    }

    public static class CollectionLinkFormats
    {
        // In the order they are offered.
        public static IReadOnlyList<CollectionLinkFormat> All { get; } =
            [CollectionLinkFormat.BbCode, CollectionLinkFormat.Markdown, CollectionLinkFormat.Html, CollectionLinkFormat.Plain];

        // The link to `url` with `text` as its words, as `format` writes it. A square bracket in the words is written round in
        // BBCode and Markdown: in BBCode a "[/url]" typed in a title would end the link, and in Markdown a "]" ends its words.
        // HTML encodes the words and the address, so a "<" or a quote typed in a title stays a character.
        public static string Write(CollectionLinkFormat format, string url, string text)
        {
            ArgumentNullException.ThrowIfNull(url);

            if (format == CollectionLinkFormat.Html)
                return $"<a href=\"{WebUtility.HtmlEncode(url)}\">{WebUtility.HtmlEncode(text ?? string.Empty)}</a>";

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
            CollectionLinkFormat.Html => "HTML",
            CollectionLinkFormat.Plain => "Plain link",
            _ => "BBCode"
        };

        // Where it works - the examples beside its name.
        public static string WhereOf(CollectionLinkFormat format) => format switch
        {
            CollectionLinkFormat.Markdown => "Reddit, Discord, GitHub and Discourse",
            CollectionLinkFormat.Html => "Web pages, blogs and HTML mail",
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
