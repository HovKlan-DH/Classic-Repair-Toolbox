using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Theming
{
    // ###########################################################################################
    // TEXT IN PIECES, THE MARKED ONES BOLD - a count's number, a person's name, a notice's data
    // source. A TextBlock cannot mix weights within one Text, so a line with a bold piece is built
    // of Runs, and its Text is then NULL: read one through TextOf, never through Text. With nothing
    // marked it is plain Text, which lays out cheaper. Whatever the block showed before is replaced.
    //
    // The Maintainer tab's lines (TabMaintainer.ShowRuns) and Main's notice cards (NoticeEmphasis)
    // both write through here.
    // ###########################################################################################
    internal static class BoldRuns
    {
        public static void Show(TextBlock block, IReadOnlyList<(string Text, bool IsBold)> runs)
        {
            block.Inlines?.Clear();
            block.Text = null;

            if (!runs.Any(run => run.IsBold))
            {
                block.Text = string.Concat(runs.Select(run => run.Text));
                return;
            }

            block.Inlines ??= new InlineCollection();

            foreach ((string text, bool isBold) in runs)
            {
                block.Inlines.Add(new Run(text)
                {
                    FontWeight = isBold ? FontWeight.Bold : block.FontWeight
                });
            }
        }

        // ###########################################################################################
        // A sentence that says its own bold pieces as "<b>...</b>" (2026-11-03) - so a wording class,
        // which holds no control, can say what is bold. A tag with no partner, or one inside another,
        // is read as if it were not there: this is CRT's own words, and a slip must not lose any of them.
        // ###########################################################################################
        public static IReadOnlyList<(string Text, bool IsBold)> Marked(string? text)
        {
            var runs = new List<(string Text, bool IsBold)>();
            string rest = text ?? string.Empty;

            while (rest.Length > 0)
            {
                int open = rest.IndexOf(OpenTag, StringComparison.Ordinal);
                int close = open < 0 ? -1 : rest.IndexOf(CloseTag, open + OpenTag.Length, StringComparison.Ordinal);

                if (close < 0)
                {
                    runs.Add((rest.Replace(OpenTag, string.Empty, StringComparison.Ordinal).Replace(CloseTag, string.Empty, StringComparison.Ordinal), false));
                    break;
                }

                if (open > 0)
                {
                    runs.Add((rest[..open].Replace(CloseTag, string.Empty, StringComparison.Ordinal), false));
                }

                runs.Add((rest[(open + OpenTag.Length)..close].Replace(OpenTag, string.Empty, StringComparison.Ordinal), true));
                rest = rest[(close + CloseTag.Length)..];
            }

            return runs;
        }

        private const string OpenTag = "<b>";
        private const string CloseTag = "</b>";

        public static void ShowMarked(TextBlock block, string? text) => BoldRuns.Show(block, BoldRuns.Marked(text));

        // The same sentence with its tags gone - what is read on screen.
        public static string Plain(string? text) => string.Concat(BoldRuns.Marked(text).Select(run => run.Text));

        // What a block says, whether it holds plain Text or bold pieces in Inlines.
        public static string TextOf(TextBlock block) =>
            block.Text ?? string.Concat(block.Inlines?.OfType<Run>().Select(run => run.Text) ?? []);

        // The bold pieces a block shows, in order - none for plain Text.
        public static IReadOnlyList<string> BoldPiecesOf(TextBlock block) =>
            block.Inlines?.OfType<Run>()
                .Where(run => run.FontWeight >= FontWeight.Bold)
                .Select(run => run.Text ?? string.Empty)
                .ToList() ?? [];
    }
}
