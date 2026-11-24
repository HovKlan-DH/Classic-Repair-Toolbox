using Avalonia.Controls;

namespace Handlers.Theming
{
    // ###########################################################################################
    // A LIST SCROLLED TO A ROW AFTER ITS ROWS CHANGED (2026-11-24, measured on Avalonia 12.1.6 and 12.2.0 alike, found
    // by the move to 12.2.0). Avalonia's VirtualizingStackPanel - every ListBox's - cannot scroll to a row put in since
    // the last layout pass: ScrollIntoView builds the row, the panel's own layout then lets go of it as lying outside the
    // old view, and nothing moves (399 rows put in above the selected one: offset 0, before and after a render tick).
    // Laid out first, it gets there. So a list is laid out first when its rows changed since - only then, as a layout
    // pass is the whole window's.
    //
    // And never a hidden list: 12.2.0 keeps a scroll asked of a list not on screen and does it as the list is next
    // shown, where 12.1.6 did nothing - and the lists on the far left are hidden and shown as the "Contribute" tab is
    // left and entered, each put where the other stood (Main.ListPosition), which a scroll left waiting would undo.
    // ###########################################################################################
    internal static class ListScrolling
    {
        public static void ToRow(ListBox list, int index)
        {
            if (index < 0 || !list.IsEffectivelyVisible)
            {
                return;
            }

            if (list.ItemsPanelRoot is { IsMeasureValid: false } or { IsArrangeValid: false })
            {
                list.UpdateLayout();
            }

            list.ScrollIntoView(index);
        }
    }
}
