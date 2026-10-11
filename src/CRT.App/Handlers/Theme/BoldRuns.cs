using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
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
