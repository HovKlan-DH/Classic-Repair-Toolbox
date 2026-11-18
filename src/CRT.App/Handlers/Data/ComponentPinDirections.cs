using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A COMPONENT'S PINS AS THE "Schematics" TAB SAYS THEM ON THE SCHEMATIC IMAGE (owner requests, 2026-11-17): the direction
    // of the pin under the pointer ("Direction Input", a badge of its own), and each pin's color when they are numbered with
    // SHIFT held - THE DIRECTION THE PINOUT GIVES IT (PinDirectionKinds.ByPin's order): what the component's pin list says in
    // the region shown, else the datasheet's of a chip CRT knows (LogicChips.For). A pad KiCad numbers is matched by its
    // number as every reading of a pin goes (ComponentVocabulary.PinNumberOf - "2 [64K]" is pin 2), a pad named otherwise
    // ("A1") by its name.
    // ###########################################################################################
    public sealed class ComponentPinDirections
    {
        private readonly List<ComponentPinEntry> thisPins;

        private ComponentPinDirections(List<ComponentPinEntry> pins, LogicChip? chip)
        {
            this.thisPins = pins;
            this.Chip = chip;
        }

        // The chip CRT holds the datasheet of for this component - null for none.
        public LogicChip? Chip { get; }

        // No component, no pins: every pin without a direction.
        public static readonly ComponentPinDirections None = new([], null);

        // The component `label` of `board` in `region`, as the component popup reads it.
        public static ComponentPinDirections For(BoardData? board, string? label, string? region)
        {
            if (board is null || string.IsNullOrWhiteSpace(label))
            {
                return ComponentPinDirections.None;
            }

            List<ComponentPinEntry> pins = ComponentPopupContent.PinEntriesOf(board.ComponentPins, label, region);
            ComponentEntry? entry = ComponentPaneEdits.RowShown(board, label, region);

            return new ComponentPinDirections(pins, entry is null ? null : LogicChips.For(entry, pins));
        }

        // ###########################################################################################
        // The direction of `pad` in the pin list's words ("Input", "Output", "Power", "Unknown", ...) - "" when neither the
        // list nor a datasheet says one.
        // ###########################################################################################
        public string WordOf(string? pad)
        {
            string wanted = (pad ?? string.Empty).Trim();

            if (wanted.Length == 0)
            {
                return string.Empty;
            }

            int? number = ComponentVocabulary.PinNumberOf(wanted);

            ComponentPinEntry? row = this.thisPins.FirstOrDefault(pin =>
                !string.IsNullOrWhiteSpace(pin.Direction) &&
                (number is int n
                    ? ComponentVocabulary.PinNumberOf(pin.Pin) == n
                    : string.Equals((pin.Pin ?? string.Empty).Trim(), wanted, StringComparison.OrdinalIgnoreCase)));

            // The pinout's own rule (PinDirectionKinds.WordOf) - pins 1 to the chip's last only: a datasheet's pin 0 is no pin
            // but the wires inside a unit (code review, 2026-11-18: a pad numbered "0" was given a terminal inside the chip's).
            return PinDirectionKinds.WordOf(row?.Direction, number, this.Chip);
        }

        // The color a numbered pin is drawn in - the pinout's for that direction.
        public PinDirectionKind KindOf(string? pad) => PinDirectionKinds.Of(this.WordOf(pad));
    }
}
