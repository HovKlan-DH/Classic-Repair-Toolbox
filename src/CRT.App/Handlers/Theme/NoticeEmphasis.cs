using Avalonia.Controls;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Handlers.Theming
{
    // ###########################################################################################
    // WHAT A NOTICE CARD PUTS IN BOLD (owner request, 2026-10-11: "Checking data from the <b>stable</b>
    // source", "[<b>3</b>] files updated in background", "and do that for others also"). Every card
    // in Main's NoticeStack whose words are set in code writes them through Show, so all of them
    // follow the same two rules:
    //
    //   - A value in square brackets is bold, the brackets not: "[3] files updated",
    //     "Version [3.0.0] is available", "Downloading file [2] of [9]".
    //   - The data source's NAME is bold where a notice names a source: "stable" in "the stable
    //     source", "BETA" in "the BETA source" (AppConfig.GetOnlineSourceLabel's two names).
    //
    // NOTHING INSIDE DOUBLE QUOTES IS TOUCHED: a quoted piece is a control's label ("tick "Download
    // data from the BETA source instead of the stable source""), and it reads as it does on the
    // control.
    //
    // A card's words are then in Inlines and its Text is NULL (BoldRuns) - read them through TextOf.
    // Runs is pure, so the rules are tested without a window (NoticeEmphasisTests); MainWindowTests
    // checks that every card on screen shows them.
    // ###########################################################################################
    internal static class NoticeEmphasis
    {
        // A quoted label, left as it is - an opening quote with no closing one quotes nothing.
        private static readonly Regex Quoted = new("\"[^\"]*\"", RegexOptions.CultureInvariant);

        // A bracketed value (one line, no quote in it), or a source's name before " source".
        private static readonly Regex Emphasised = new(
            @"\[(?<value>[^\[\]\r\n""]+)\]|\b(?<source>stable|BETA)(?= source\b)",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        // ###########################################################################################
        // The words in pieces, the bold ones marked - neighbouring plain pieces joined. Joined
        // again, the pieces are the words exactly.
        // ###########################################################################################
        public static IReadOnlyList<(string Text, bool IsBold)> Runs(string? text)
        {
            var runs = new List<(string Text, bool IsBold)>();

            if (string.IsNullOrEmpty(text))
            {
                return runs;
            }

            int at = 0;

            foreach (Match quote in NoticeEmphasis.Quoted.Matches(text))
            {
                NoticeEmphasis.AddUnquoted(runs, text[at..quote.Index]);
                NoticeEmphasis.Add(runs, quote.Value, isBold: false);
                at = quote.Index + quote.Length;
            }

            NoticeEmphasis.AddUnquoted(runs, text[at..]);
            return runs;
        }

        // Puts the words on a card's TextBlock, with their bold pieces.
        public static void Show(TextBlock block, string? text) => BoldRuns.Show(block, NoticeEmphasis.Runs(text));

        // What a card's TextBlock says, bold pieces or not.
        public static string TextOf(TextBlock block) => BoldRuns.TextOf(block);

        private static void AddUnquoted(List<(string Text, bool IsBold)> runs, string text)
        {
            int at = 0;

            foreach (Match match in NoticeEmphasis.Emphasised.Matches(text))
            {
                if (match.Groups["value"].Success)
                {
                    NoticeEmphasis.Add(runs, text[at..match.Index] + "[", isBold: false);
                    NoticeEmphasis.Add(runs, match.Groups["value"].Value, isBold: true);
                    NoticeEmphasis.Add(runs, "]", isBold: false);
                }
                else
                {
                    NoticeEmphasis.Add(runs, text[at..match.Index], isBold: false);
                    NoticeEmphasis.Add(runs, match.Value, isBold: true);
                }

                at = match.Index + match.Length;
            }

            NoticeEmphasis.Add(runs, text[at..], isBold: false);
        }

        private static void Add(List<(string Text, bool IsBold)> runs, string text, bool isBold)
        {
            if (text.Length == 0)
            {
                return;
            }

            if (!isBold && runs.Count > 0 && !runs[^1].IsBold)
            {
                runs[^1] = (runs[^1].Text + text, false);
                return;
            }

            runs.Add((text, isBold));
        }
    }
}
