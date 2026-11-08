using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A CHANGELOG VERSION'S HEADING AS THE "ABOUT" TAB SHOWS IT (owner request, 2026-11-06: "now it shows each
    // version header like this: <b>3.0.0, 2026-October-11</b> - Please show it like this: Version <b>3.0.0</b>,
    // 2026-October-11"). CHANGELOG.md writes each version "# 3.0.0, 2026-October-11" - its words are the owner's,
    // read and never written - so the heading is reworded where it is shown: "Version " before it, only the
    // number bold.
    //
    // Only a heading that is a version and a date and nothing else: one with any styling of its own, or another
    // shape ("Known issues"), is left as GitHub draws it.
    // ###########################################################################################
    internal static class ChangelogVersionHeadings
    {
        public const string VersionWord = "Version ";

        // "4.0.0, 2026-November-?", "3.1.1, 2026-November-5", "4.0.0-beta.2, ..." - the number, then the rest.
        private static readonly Regex VersionAndDate = new(@"^(\d+\.\d+\.\d+[^,\s]*)(, .+)$", RegexOptions.CultureInvariant);

        // The blocks with every version's heading reworded, the rest as they were.
        public static IReadOnlyList<MarkdownBlock> Of(IReadOnlyList<MarkdownBlock> blocks) =>
            blocks.Select(block => block is MarkdownHeading heading ? Reworded(heading) : block).ToList();

        internal static MarkdownHeading Reworded(MarkdownHeading heading)
        {
            if (heading.BoldAsMarked || heading.Spans.Any(span => span.Style != MarkdownStyle.None || span.Url != null))
            {
                return heading;
            }

            Match match = VersionAndDate.Match(string.Concat(heading.Spans.Select(span => span.Text)).Trim());
            if (!match.Success)
            {
                return heading;
            }

            return new MarkdownHeading(
                heading.Level,
                [
                    new MarkdownSpan(VersionWord, MarkdownStyle.None, null),
                    new MarkdownSpan(match.Groups[1].Value, MarkdownStyle.Bold, null),
                    new MarkdownSpan(match.Groups[2].Value, MarkdownStyle.None, null),
                ],
                BoldAsMarked: true);
        }
    }
}
