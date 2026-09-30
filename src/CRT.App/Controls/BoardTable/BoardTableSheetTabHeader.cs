using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using System.Globalization;

namespace CRT
{
    // ###########################################################################################
    // WHAT A SHEET TAB SAYS: its name with the change count - "Components (3)" - and, when rows on
    // it are flagged, a violet pill with their number: "[ 6 flagged ]" (owner request, 2026-09-26:
    // "the flagged counters should get visualized also in the tabs headline, as these will be
    // important to address").
    //
    // *** THE FLAGGED NUMBER IS NOT IN THE BRACKETS. *** The bracketed count is the sheet's changes,
    // which must keep agreeing with the "N rows changed" BoardDataDiffer counts above the table - a
    // duplicate or an incomplete row is not a change. So flagged rows get their own pill, in their
    // own colour (BoardTable_Flagged_Bg, the colour key's "Flagged" and the rows' own violet).
    //
    // A record, so a tab whose counts did not move keeps its header; its ToString is the text.
    // ###########################################################################################
    internal sealed record BoardTableSheetTabHeader(string Text, int Flagged)
    {
        public override string ToString() => this.Text;

        public string? FlaggedText =>
            this.Flagged > 0 ? $"{this.Flagged.ToString(CultureInfo.InvariantCulture)} flagged" : null;

        public static IDataTemplate Template { get; } = new FuncDataTemplate<BoardTableSheetTabHeader>(
            (header, _) =>
            {
                var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                panel.Children.Add(new TextBlock { Text = header?.Text, VerticalAlignment = VerticalAlignment.Center });

                if (header?.FlaggedText is string flagged)
                {
                    panel.Children.Add(new Border
                    {
                        Classes = { "SheetTabFlagged" },
                        Child = new TextBlock { Text = flagged }
                    });
                }

                return panel;
            });
    }
}
