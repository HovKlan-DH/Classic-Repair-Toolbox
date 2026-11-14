using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // One pad of the board's KiCad data on a net: the footprint's reference, the pad's number and the net's name.
    internal readonly record struct KiCadPadOnNet(string Reference, string Pad, string Net);

    // ###########################################################################################
    // One line under "Connected to:", in the pieces it is drawn in: "[", the board label in bold, the rest of the
    // component's name and "] pin ", the pin number in bold, and with several nets highlighted " [PHI2]" (owner requests,
    // 2026-11-13: "[ <b>U17</b> | 74LS373 ] pin <b>5</b>", then no space inside the brackets, "a little more condensed").
    // Pieces, never "<b>" in the text: a friendly name is typed by people, and BoldRuns would read a "<b>" in it as its own.
    // ###########################################################################################
    public sealed record KiCadNetConnectionLine(string Label, string Rest, string Pin, string Net)
    {
        // Between the label and the pin: " | 74LS373] pin ", or "] pin " for a component with nothing else filled.
        public string AfterLabel => $"{this.Rest}] pin ";

        // After the pin: " [PHI2]" with several nets highlighted, else nothing.
        public string AfterPin => this.Net.Length > 0 ? $" [{this.Net}]" : string.Empty;

        // The line as read, the bold aside.
        public string Text => $"[{this.Label}{this.AfterLabel}{this.Pin}{this.AfterPin}";
    }

    // ###########################################################################################
    // WHAT THE "KiCad nets" PANEL OF THE "Schematics" TAB LISTS UNDER "Connected to:" - every pad on the nets
    // highlighted, one line each.
    //
    //   - EACH COMPONENT IS NAMED AS ITS COMPONENT INFO WINDOW'S TITLE NAMES IT (owner request, 2026-11-13: "Could it
    //     instead show the "display name", which is identical to the window title for the component info window"):
    //     ComponentPopupContent.TitleOf, of the component's row in the region shown, as the window picks it
    //     (ComponentImageQueries.PickComponentEntry) - "[U5 | CIA 1 | 6581] pin 13", the label and the pin bold (owner
    //     requests, the same day). A reference the board has no component of is its reference alone, "[U5] pin 13".
    //   - With several nets highlighted each line ends in its net: "[U5 | CIA 1 | 6581] pin 13 [PHI2]".
    //   - In the board's component order, then pin number, then the line; each line once.
    //
    // PURE: the pads, the board's components and the region in, the lines out.
    // ###########################################################################################
    internal static class KiCadNetConnections
    {
        public static List<KiCadNetConnectionLine> Lines(
            IEnumerable<KiCadPadOnNet> pads,
            IReadOnlyCollection<string> activeNets,
            IReadOnlyList<ComponentEntry>? components,
            string? region)
        {
            ArgumentNullException.ThrowIfNull(pads);
            ArgumentNullException.ThrowIfNull(activeNets);

            var nets = new HashSet<string>(activeNets, StringComparer.OrdinalIgnoreCase);
            string regionShown = region?.Trim() ?? string.Empty;

            // The board's order and each label's rows, made once.
            var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var rowsOf = new Dictionary<string, List<ComponentEntry>>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < (components?.Count ?? 0); index++)
            {
                ComponentEntry row = components![index];
                string label = row.BoardLabel?.Trim() ?? string.Empty;

                if (label.Length == 0)
                {
                    continue;
                }

                order.TryAdd(label, index);

                if (!rowsOf.TryGetValue(label, out List<ComponentEntry>? rows))
                {
                    rowsOf[label] = rows = [];
                }

                rows.Add(row);
            }

            // A component's name once per reference, not once per pad: its label as the board spells it, and the rest of
            // the window's title after it - " | CIA 1 | 6581", or "" with nothing else filled.
            var names = new Dictionary<string, (string Label, string After)>(StringComparer.OrdinalIgnoreCase);

            (string Label, string After) NameOf(string reference)
            {
                if (names.TryGetValue(reference, out var known))
                {
                    return known;
                }

                (string Label, string After) name = (reference, string.Empty);

                if (rowsOf.TryGetValue(reference, out List<ComponentEntry>? rows))
                {
                    string label = rows[0].BoardLabel!.Trim();
                    string rest = ComponentPopupContent.TitleOf(null, ComponentImageQueries.PickComponentEntry(rows, regionShown));

                    name = (label, rest.Length > 0 ? $" | {rest}" : string.Empty);
                }

                names[reference] = name;
                return name;
            }

            var lines = new List<(KiCadNetConnectionLine Line, int Order, int Pad)>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (KiCadPadOnNet pad in pads)
            {
                string reference = pad.Reference?.Trim() ?? string.Empty;
                string net = pad.Net?.Trim() ?? string.Empty;

                if (reference.Length == 0 || !nets.Contains(net))
                {
                    continue;
                }

                string number = pad.Pad?.Trim() is { Length: > 0 } padNumber ? padNumber : "?";
                (string label, string rest) = NameOf(reference);
                var line = new KiCadNetConnectionLine(label, rest, number, nets.Count > 1 ? net : string.Empty);

                if (!seen.Add(line.Text))
                {
                    continue;
                }

                lines.Add((
                    line,
                    order.TryGetValue(reference, out int place) ? place : int.MaxValue,
                    int.TryParse(number, NumberStyles.Integer, CultureInfo.InvariantCulture, out int padIndex) ? padIndex : int.MaxValue));
            }

            return lines
                .OrderBy(item => item.Order)
                .ThenBy(item => item.Pad)
                .ThenBy(item => item.Line.Text, StringComparer.OrdinalIgnoreCase)
                .Select(item => item.Line)
                .ToList();
        }
    }
}
