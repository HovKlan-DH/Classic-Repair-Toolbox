using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A BOARD'S REGION BUTTONS PUT IN THEIR GRID (2026-11-09: a board may come in regions of its own) - Main's under the
    // component list, every place RegionButtonLayout's. The grid's columns and rows are its markup's and NEVER change (owner
    // report, 2026-11-09: CRT threw at launch - RegionButtonLayout's header says why): a button only takes its row, column,
    // span and margin.
    //
    // WHICH BUTTON EACH REGION IS (ButtonsFor) is one rule for Main's buttons and the component popup's (code review,
    // 2026-11-09: written twice, the two had drifted - the popup's own buttons cut a long name off).
    // ###########################################################################################
    internal static class RegionButtonGrid
    {
        // Each of `buttons` in its place; a button on the second row 6px under the first.
        public static void Place(IReadOnlyList<Button> buttons)
        {
            for (int index = 0; index < buttons.Count; index++)
            {
                (int row, int column, int span) = RegionButtonLayout.PlaceOf(index, buttons.Count);
                Grid.SetRow(buttons[index], row);
                Grid.SetColumn(buttons[index], column);
                Grid.SetColumnSpan(buttons[index], span);
                buttons[index].Margin = new Thickness(0, row > 0 ? RegionButtonLayout.Gap : 0, 0, 0);
            }
        }

        // A region of the board's own in its color: class Own1 to Own4 (RegionChoice), the others taken off.
        public static void UseOwnColors(Button button, int? number)
        {
            for (int each = 1; each <= 4; each++)
            {
                button.Classes.Set($"Own{each}", number == each);
            }
        }

        // ###########################################################################################
        // The buttons for `regions`, in the board's order: "PAL" and "NTSC" are `pal` and `ntsc`, the buttons of the markup; a
        // region of the board's own takes the next of `own`, `make` making one - added to `host` - when there are too few. Each
        // button's Tag is its region, an own one's colors RegionChoice's; every button not taken is hidden. Where each stands
        // is the caller's.
        // ###########################################################################################
        public static List<Button> ButtonsFor(IReadOnlyList<string> regions, Button pal, Button ntsc, List<Button> own, Func<Button> make, Panel host)
        {
            ArgumentNullException.ThrowIfNull(regions);
            ArgumentNullException.ThrowIfNull(own);

            var buttons = new List<Button>();
            int next = 0;

            pal.Tag = BoardRegions.Pal;
            ntsc.Tag = BoardRegions.Ntsc;

            foreach (string region in regions)
            {
                if (string.Equals(region, BoardRegions.Pal, StringComparison.OrdinalIgnoreCase))
                {
                    buttons.Add(pal);
                    continue;
                }

                if (string.Equals(region, BoardRegions.Ntsc, StringComparison.OrdinalIgnoreCase))
                {
                    buttons.Add(ntsc);
                    continue;
                }

                while (own.Count <= next)
                {
                    Button made = make();
                    made.Classes.Add("RegionToggle");
                    own.Add(made);
                    host.Children.Add(made);
                }

                Button button = own[next++];
                button.Tag = region;
                RegionButtonGrid.UseOwnColors(button, RegionChoice.OwnNumberOf(regions, region));
                buttons.Add(button);
            }

            foreach (Button button in own.Append(pal).Append(ntsc))
            {
                button.IsVisible = buttons.Contains(button);
            }

            return buttons;
        }

        // A block for an own region's name that wraps rather than being cut (the render, 2026-11-09) - made once with its button.
        public static TextBlock NameBlock() => new() { TextWrapping = TextWrapping.Wrap, TextAlignment = TextAlignment.Center };

        // `text` on `button` - in its name block when it has one, set only where it differs, so nothing is made again.
        public static void ShowText(Button button, string text)
        {
            if (button.Content is TextBlock block)
            {
                if (!string.Equals(block.Text, text, StringComparison.Ordinal))
                {
                    block.Text = text;
                }
            }
            else if (!Equals(button.Content, text))
            {
                button.Content = text;
            }
        }

        // What `button` says, in its name block or as its content.
        public static string TextOf(Button button) => button.Content as string ?? (button.Content as TextBlock)?.Text ?? string.Empty;
    }
}
