using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // One pad of a footprint, where it lies on the board (millimetres) - enough to tell a wide DIP
    // from a narrow one.
    public sealed record KiCadPadPlace(string Number, double X, double Y);

    // A footprint of the board's PCB file: its reference, its value, the footprint it is drawn with
    // ("Package_DIP:DIP-40_W15.24mm"), its datasheet property and its pads.
    public sealed record KiCadFootprintInfo(string Reference, string Value, string Footprint, string Datasheet, IReadOnlyList<KiCadPadPlace> Pads);

    // A symbol of the board's schematic: its reference, its value and its datasheet property.
    public sealed record KiCadSymbolInfo(string Reference, string Value, string Datasheet);

    // A component the board's KiCad data has, as CRT would add it: its label, its value, the category
    // its reference suggests, its package ("" when the footprint does not say) and its datasheet's
    // address ("" for none).
    public sealed record KiCadBoardComponent(string Label, string Value, string Category, string Package, string Datasheet);

    // What adding them did: how many components, how many with their pins, how many with a link.
    public sealed record KiCadComponentsAdded(int Components, int WithPins, int WithDatasheet);

    // ###########################################################################################
    // A NEW BOARD'S COMPONENTS FROM ITS KICAD DATA (CRT 4.0 step 3 - the plan: "Components: one press
    // adds every KiCad component that isn't on the board yet: label and value from KiCad, category
    // guessed from the reference prefix, package from the footprint name (a DIP-40 footprint gives a
    // pin list already filled with numbers and nets), a datasheet link from the KiCad symbol").
    //
    // Plain values in, like KiCadComponentPins: the KiCad types live in CRT.App (KiCadComponentTable).
    //
    //   - THE COMPONENTS ARE THE FIRST PCB FILE'S FOOTPRINTS (KiCadReferenceMatcher's rule - only it
    //     lights up copper), once per reference; a project with no PCB file, its schematic's symbols.
    //     Not a component: a placeholder reference ("REF**", "#PWR01", "U?"), a footprint with no
    //     numbered pad (a logo, a graphic), a mounting hole, a fiducial, a net tie.
    //   - THE CATEGORY is the one the live boards use for that reference prefix (measured on the
    //     converted stable tree, 2026-10-26: every C is a Capacitor, every R and RP a Resistor, U and
    //     IC an IC, CR, D and DP a Diode, CN and P a Connector, Q and TR a Transistor) - KiCad's own
    //     convention where the boards have none. The rest is Misc, as ferrite beads, filters, coils,
    //     crystals, switches and jumpers are there. J IS A JUMPER, Misc (owner decision, 2026-10-31:
    //     "the 'J' component should be a 'Misc' category, as it is a jumper") - on the live boards 30
    //     of 32 J parts are jumpers filed as Misc, and every J of the shipped KiCad projects is one;
    //     KiCad's own libraries call a connector J, which a contributor then changes by hand.
    //   - THE PACKAGE is ComponentVocabulary's word when the footprint's name says it AND its pads
    //     agree: a name STARTING "DIP-14_W7.62mm", "DIP16" or "PDIP-40" with that many pads
    //     ("Oscillator_DIP-14" is a can, "SOIC-DIP-8" surface-mounted), a 24- or 28-pin DIP "wide" when
    //     its rows are 600 mil apart; SIP and a resistor array; TO-92 and TO-220 with three pads; and
    //     since 2026-11-03 (owner request: "ideally guess their 'package'"), by what the footprints of the
    //     six published boards with a PCB file call them: DO-35, DO-41, HC-49, HC-18, a pin header, a
    //     three-legged EMI filter, and with two pads Axial, Radial or Disc as the name says - a resistor,
    //     a ferrite bead, an inductor, a diode by theirs, and a resistor whatever its footprint is called
    //     (the Spectrum's "Passive_small": no surface-mounted one on any of them, and "SMD" or "0805" in
    //     a name, or "SMD" in its library, says otherwise). Anything else is left empty - a "CAP_P5.08mm" says only its pitch, and
    //     "Other" says nothing more.
    //   - THE DATASHEET is the symbol's, else the footprint's, when it is a web address (TextLinkFinder's
    //     shapes - "www." opened as "https://"); KiCad's "~" and a library path are none.
    //   - THE VALUE is KiCad's as it stands; "~" is none.
    //   - THE CATEGORIES ARE ADDED IN ALPHABETICAL ORDER (owner request, 2026-11-03; the chips first until
    //     then) - the board's own stay where they are, only a new one is placed among them. The order is
    //     then the contributor's to drag.
    // ###########################################################################################
    public static partial class KiCadBoardComponents
    {
        // ###########################################################################################
        // Every component of the KiCad data, once each, in the order a new board lists them: by
        // category alphabetically, then in natural label order.
        // ###########################################################################################
        public static List<KiCadBoardComponent> Of(IReadOnlyList<KiCadFootprintInfo> footprints, IReadOnlyList<KiCadSymbolInfo> symbols)
        {
            ArgumentNullException.ThrowIfNull(footprints);
            ArgumentNullException.ThrowIfNull(symbols);

            var symbolsByReference = new Dictionary<string, KiCadSymbolInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (KiCadSymbolInfo symbol in symbols.Where(symbol => KiCadBoardComponents.IsComponentReference(symbol.Reference)))
                symbolsByReference.TryAdd(symbol.Reference.Trim(), symbol);

            var found = new Dictionary<string, KiCadBoardComponent>(StringComparer.OrdinalIgnoreCase);

            if (footprints.Count > 0)
            {
                foreach (KiCadFootprintInfo footprint in footprints.Where(KiCadBoardComponents.IsComponent))
                {
                    string label = footprint.Reference.Trim();

                    if (found.ContainsKey(label))
                        continue;

                    symbolsByReference.TryGetValue(label, out KiCadSymbolInfo? symbol);

                    string value = KiCadBoardComponents.ValueOf(footprint.Value) is { Length: > 0 } footprintValue
                        ? footprintValue
                        : KiCadBoardComponents.ValueOf(symbol?.Value);

                    string datasheet = KiCadBoardComponents.LinkOf(symbol?.Datasheet) is { Length: > 0 } symbolLink
                        ? symbolLink
                        : KiCadBoardComponents.LinkOf(footprint.Datasheet);

                    found[label] = new KiCadBoardComponent(
                        label,
                        value,
                        KiCadBoardComponents.CategoryOf(label),
                        KiCadBoardComponents.PackageOf(footprint.Footprint, footprint.Pads, label),
                        datasheet);
                }
            }
            else
            {
                foreach ((string label, KiCadSymbolInfo symbol) in symbolsByReference)
                {
                    found[label] = new KiCadBoardComponent(
                        label,
                        KiCadBoardComponents.ValueOf(symbol.Value),
                        KiCadBoardComponents.CategoryOf(label),
                        string.Empty,
                        KiCadBoardComponents.LinkOf(symbol.Datasheet));
                }
            }

            return found.Values
                .OrderBy(component => component.Category, StringComparer.OrdinalIgnoreCase)
                .ThenBy(component => component.Label, NaturalLabelComparer.Instance)
                .ToList();
        }

        // ###########################################################################################
        // WHAT KICAD KNOWS BUT CANNOT PLACE (owner request, 2026-11-01: "components known by KiCad but
        // needs manual marking"): the references of the schematic's symbols no footprint of the PCB
        // file has - there is no footprint to draw a rectangle from. In natural order. With no PCB file
        // at all, every symbol.
        // ###########################################################################################
        public static List<string> OnlyInSchematic(IReadOnlyList<KiCadFootprintInfo> footprints, IReadOnlyList<KiCadSymbolInfo> symbols)
        {
            ArgumentNullException.ThrowIfNull(footprints);
            ArgumentNullException.ThrowIfNull(symbols);

            var placed = new HashSet<string>(footprints.Select(footprint => footprint.Reference?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);

            return symbols
                .Where(symbol => KiCadBoardComponents.IsComponentReference(symbol.Reference))
                .Select(symbol => symbol.Reference.Trim())
                .Where(reference => !placed.Contains(reference))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(reference => reference, NaturalLabelComparer.Instance)
                .ToList();
        }

        // The ones `board` has no component of - by label, trimmed, capitals ignored, in any region.
        public static List<KiCadBoardComponent> NotOn(BoardData board, IEnumerable<KiCadBoardComponent> components)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(components);

            var labels = new HashSet<string>(board.Components.Select(row => row.BoardLabel?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);

            return components.Where(component => !labels.Contains(component.Label)).ToList();
        }

        // ###########################################################################################
        // `board` with every one of `components` it lacks added - each placed where a new component
        // goes (ComponentPlacement), with its pins when its package is known (`pinsOf` - the numbers,
        // names and directions KiCad has; the net is never stored) and a "Datasheet for <value>" link
        // when it has a datasheet. The same board when there is nothing to add.
        // ###########################################################################################
        public static BoardData WithComponents(
            BoardData board,
            IReadOnlyList<KiCadBoardComponent> components,
            Func<string, IReadOnlyList<KiCadPin>> pinsOf,
            out KiCadComponentsAdded added)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(components);
            ArgumentNullException.ThrowIfNull(pinsOf);

            List<KiCadBoardComponent> missing = KiCadBoardComponents.NotOn(board, components);

            if (missing.Count == 0)
            {
                added = new KiCadComponentsAdded(0, 0, 0);
                return board;
            }

            var rows = board.Components.ToList();
            var pins = board.ComponentPins.ToList();
            var links = board.ComponentLinks.ToList();
            int withPins = 0;
            int withDatasheet = 0;

            foreach (KiCadBoardComponent component in missing)
            {
                rows.Insert(
                    KiCadBoardComponents.IndexFor(rows, component),
                    new ComponentEntry
                    {
                        BoardLabel = component.Label,
                        TechnicalNameOrValue = component.Value,
                        Category = component.Category,
                        Package = component.Package,
                    });

                // A label nothing else lists yet - pin and link rows left over under it are kept as
                // they are, never doubled.
                if (component.Package.Length > 0 && !pins.Any(row => KiCadBoardComponents.Same(row.BoardLabel, component.Label)))
                {
                    List<ComponentPinEntry> componentPins = pinsOf(component.Label)
                        .Select(pin => new ComponentPinEntry { BoardLabel = component.Label, Pin = pin.Pin, Name = pin.Name, Direction = pin.Direction })
                        .ToList();

                    if (componentPins.Count > 0)
                    {
                        pins = ComponentPlacement.ReplaceRowsForLabel(pins, component.Label, row => row.BoardLabel, componentPins);
                        withPins++;
                    }
                }

                if (component.Datasheet.Length > 0 && !links.Any(row => KiCadBoardComponents.Same(row.BoardLabel, component.Label)))
                {
                    var link = new ComponentLinkEntry
                    {
                        BoardLabel = component.Label,
                        Name = component.Value.Length > 0 ? $"Datasheet for {component.Value}" : "Datasheet",
                        Url = component.Datasheet,
                    };

                    links = ComponentPlacement.ReplaceRowsForLabel(links, component.Label, row => row.BoardLabel, [link]);
                    withDatasheet++;
                }
            }

            added = new KiCadComponentsAdded(missing.Count, withPins, withDatasheet);

            return board.WithComponentSections(components: rows, componentLinks: links, componentPins: pins);
        }

        // ###########################################################################################
        // Where an added component's row goes: into its category as every new component does
        // (ComponentPlacement) - and a category the board does not have yet before the first one that
        // sorts after it, so the import's categories come out in alphabetical order around the board's.
        // ###########################################################################################
        private static int IndexFor(IReadOnlyList<ComponentEntry> rows, KiCadBoardComponent component)
        {
            bool hasCategory = rows.Any(row => string.Equals(row.Category?.Trim(), component.Category, StringComparison.OrdinalIgnoreCase));

            if (hasCategory)
            {
                return ComponentPlacement.InsertionIndex(rows, component.Category, component.Label);
            }

            for (int index = 0; index < rows.Count; index++)
            {
                if (string.Compare(rows[index].Category?.Trim(), component.Category, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    return index;
                }
            }

            return rows.Count;
        }

        // ---- the guesses --------------------------------------------------------------------------

        // ###########################################################################################
        // The category a reference's prefix suggests - its leading letters, capitals ignored.
        // ###########################################################################################
        public static string CategoryOf(string? reference)
        {
            string prefix = new((reference?.Trim() ?? string.Empty).TakeWhile(char.IsAsciiLetter).ToArray());
            prefix = prefix.ToUpperInvariant();

            return prefix switch
            {
                "C" or "CT" or "CV" or "CP" => "Capacitor",
                "R" or "RP" or "RN" or "RA" or "RV" => "Resistor",
                "IC" => "IC",
                _ when prefix.StartsWith('U') => "IC",
                "CR" or "D" or "DA" or "DP" or "LED" or "ZD" => "Diode",
                "P" or "KB" or "CON" or "EAR" or "MIC" => "Connector",
                _ when prefix.StartsWith("CN", StringComparison.Ordinal) => "Connector",
                "Q" or "TR" => "Transistor",
                _ => "Misc",
            };
        }

        // ###########################################################################################
        // The package a footprint's name says, when its pads agree - see the class comment. `reference`:
        // the component's label, which is all a resistor's plain footprint has to say.
        // ###########################################################################################
        public static string PackageOf(string? footprint, IReadOnlyList<KiCadPadPlace> pads, string? reference = null)
        {
            ArgumentNullException.ThrowIfNull(pads);

            string name = footprint?.Trim() ?? string.Empty;
            int colon = name.LastIndexOf(':');
            string library = colon >= 0 ? name[..colon] : string.Empty;

            if (colon >= 0)
                name = name[(colon + 1)..];

            // ###########################################################################################
            // No leads to speak of, by its name or the library it comes from (code review, 2026-11-04):
            // "Resistor_SMD:R_Shunt_Ohmite_LVK12" says so only there, and read as "Axial". The library only by
            // the word "SMD": a project's library can be called anything, and "C64_2010" is no chip size.
            // ###########################################################################################
            bool surfaceMounted = KiCadBoardComponents.SurfaceMounted().IsMatch(name) || KiCadBoardComponents.SurfaceMountLibrary().IsMatch(library);

            List<KiCadPadPlace> numbered = pads
                .Where(pad => (pad.Number?.Trim() ?? string.Empty).Length > 0)
                .GroupBy(pad => pad.Number.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

            Match dip = KiCadBoardComponents.Dip().Match(name);

            if (dip.Success)
            {
                // A run of digits too long for a count is none (code review, 2026-10-26: it threw).
                if (!int.TryParse(dip.Groups[1].Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count) ||
                    count % 2 != 0 ||
                    count != numbered.Count)
                    return string.Empty;

                bool wide = (count is 24 or 28) && KiCadBoardComponents.IsWide(name, numbered, count);

                return wide ? $"DIP-{count} wide" : $"DIP-{count}";
            }

            // "R_Array_Concave_4x0603" is an array too, on a surface (code review, 2026-11-04).
            if (!surfaceMounted && (KiCadBoardComponents.Sip().IsMatch(name) || KiCadBoardComponents.ResistorArray().IsMatch(name)))
                return "SIP";

            // A case CRT draws with three pins, only for a footprint with three (code review,
            // 2026-10-26: a five-pin TO-220 was drawn with three beside its five pins).
            string threePins = KiCadBoardComponents.To220().IsMatch(name) ? "TO-220"
                : KiCadBoardComponents.To92().IsMatch(name) ? "TO-92"
                : string.Empty;

            if (threePins.Length > 0)
                return numbered.Count == ComponentVocabulary.PinCountOf(threePins) ? threePins : string.Empty;

            if (KiCadBoardComponents.PinHeader().IsMatch(name))
                return "Pin header";

            return numbered.Count switch
            {
                2 => KiCadBoardComponents.TwoLeggedOf(name, reference, surfaceMounted),
                3 when KiCadBoardComponents.EmiFilter().IsMatch(name) => "EMI filter",
                _ when KiCadBoardComponents.Crystal49().IsMatch(name) => "HC-49",
                _ when KiCadBoardComponents.Crystal18().IsMatch(name) => "HC-18",
                _ => string.Empty,
            };
        }

        // What a part on two pads is, by its footprint's name - and a resistor by its reference.
        private static string TwoLeggedOf(string name, string? reference, bool surfaceMounted)
        {
            if (KiCadBoardComponents.Do35().IsMatch(name))
                return "DO-35";

            if (KiCadBoardComponents.Do41().IsMatch(name))
                return "DO-41";

            if (KiCadBoardComponents.Crystal49().IsMatch(name))
                return "HC-49";

            if (KiCadBoardComponents.Crystal18().IsMatch(name))
                return "HC-18";

            if (KiCadBoardComponents.Axial().IsMatch(name))
                return "Axial";

            if (KiCadBoardComponents.Radial().IsMatch(name))
                return "Radial";

            if (KiCadBoardComponents.Disc().IsMatch(name))
                return "Disc";

            if (surfaceMounted)
                return string.Empty;

            string prefix = new((reference?.Trim() ?? string.Empty).TakeWhile(char.IsAsciiLetter).ToArray());

            return KiCadBoardComponents.LeadedByName().IsMatch(name) || string.Equals(prefix, "R", StringComparison.OrdinalIgnoreCase)
                ? "Axial"
                : string.Empty;
        }

        // ###########################################################################################
        // A datasheet property as an address to open - "" for KiCad's "~", a blank, or anything that is
        // no web address (a library path such as "74xx/74hc_hct74.pdf").
        // ###########################################################################################
        public static string LinkOf(string? datasheet)
        {
            string text = datasheet?.Trim() ?? string.Empty;

            if (text.Length == 0 || text.Contains(' ', StringComparison.Ordinal))
                return string.Empty;

            IReadOnlyList<TextLinkSpan> spans = TextLinkFinder.FindSpans(text);

            // The whole property is the address - not one found inside other words.
            return spans.Count == 1 && spans[0].IsLink && spans[0].Length == text.Length ? spans[0].Url! : string.Empty;
        }

        // ---- the rules ----------------------------------------------------------------------------

        public static bool IsComponent(KiCadFootprintInfo footprint)
        {
            if (!KiCadBoardComponents.IsComponentReference(footprint.Reference))
                return false;

            if (!footprint.Pads.Any(pad => (pad.Number?.Trim() ?? string.Empty).Length > 0))
                return false;

            string name = footprint.Footprint ?? string.Empty;

            // A mounting hole is what the footprint IS, by its library or its own name - never a name
            // that only mentions one (owner report, 2026-11-01: the Open128's four D-sub connectors,
            // "Connector_Dsub:DSUB-9_..._MountingHolesOffset8.20mm", were left off the board).
            string ownName = name[(name.LastIndexOf(':') + 1)..];

            return !name.StartsWith("MountingHole", StringComparison.OrdinalIgnoreCase) &&
                   !ownName.StartsWith("MountingHole", StringComparison.OrdinalIgnoreCase) &&
                   !name.Contains("Fiducial", StringComparison.OrdinalIgnoreCase) &&
                   !name.Contains("NetTie", StringComparison.OrdinalIgnoreCase);
        }

        // A reference a part has - not KiCad's placeholders: "REF**", a power symbol's "#PWR01", an
        // unannotated "U?".
        public static bool IsComponentReference(string? reference)
        {
            string text = reference?.Trim() ?? string.Empty;

            return text.Length > 0 &&
                   !text.StartsWith('#') &&
                   !text.Contains('*', StringComparison.Ordinal) &&
                   !text.Contains('?', StringComparison.Ordinal);
        }

        private static string ValueOf(string? value)
        {
            string text = value?.Trim() ?? string.Empty;
            return text == "~" ? string.Empty : text;
        }

        // ###########################################################################################
        // Whether a 24- or 28-pin DIP is the 600 mil kind: its name says W15.24mm, or - for a footprint
        // whose name gives no width ("DIP24") - pin 1 and the last pin, straight across from each other,
        // lie more than 11.43 mm apart (half-way between 7.62 and 15.24).
        // ###########################################################################################
        private static bool IsWide(string name, IReadOnlyList<KiCadPadPlace> pads, int count)
        {
            if (name.Contains("W15.24", StringComparison.OrdinalIgnoreCase))
                return true;

            if (name.Contains("W7.62", StringComparison.OrdinalIgnoreCase))
                return false;

            KiCadPadPlace? first = pads.FirstOrDefault(pad => pad.Number.Trim() == "1");
            KiCadPadPlace? last = pads.FirstOrDefault(pad => pad.Number.Trim() == count.ToString(System.Globalization.CultureInfo.InvariantCulture));

            if (first is null || last is null)
                return false;

            double dx = first.X - last.X;
            double dy = first.Y - last.Y;

            return Math.Sqrt((dx * dx) + (dy * dy)) > 11.43;
        }

        private static bool Same(string? first, string second) =>
            string.Equals(first?.Trim(), second, StringComparison.OrdinalIgnoreCase);

        // "DIP-14", "DIP16", "PDIP-40", "DIP_28" - the footprint's name STARTING with it (code review,
        // 2026-10-26): "SOIC-DIP-8" is surface-mounted, "Oscillator_DIP-14" a can on a DIP's pads.
        [GeneratedRegex(@"^P?DIP[-_ ]?([0-9]+)", RegexOptions.IgnoreCase)]
        private static partial Regex Dip();

        [GeneratedRegex(@"(?<![A-Za-z])SIP", RegexOptions.IgnoreCase)]
        private static partial Regex Sip();

        [GeneratedRegex(@"(?<![A-Za-z0-9])TO-?220", RegexOptions.IgnoreCase)]
        private static partial Regex To220();

        [GeneratedRegex(@"(?<![A-Za-z0-9])TO-?92", RegexOptions.IgnoreCase)]
        private static partial Regex To92();

        // "RP_ARRAY_7WAY" - "R_Array_SIP8" says SIP itself.
        [GeneratedRegex(@"^RP?_ARRAY", RegexOptions.IgnoreCase)]
        private static partial Regex ResistorArray();

        [GeneratedRegex(@"^Pin(Header|Socket)_", RegexOptions.IgnoreCase)]
        private static partial Regex PinHeader();

        [GeneratedRegex(@"(?<![A-Za-z0-9])DO-?35(?![0-9])", RegexOptions.IgnoreCase)]
        private static partial Regex Do35();

        [GeneratedRegex(@"(?<![A-Za-z0-9])DO-?41(?![0-9])", RegexOptions.IgnoreCase)]
        private static partial Regex Do41();

        [GeneratedRegex(@"(?<![A-Za-z0-9])HC-?49", RegexOptions.IgnoreCase)]
        private static partial Regex Crystal49();

        [GeneratedRegex(@"(?<![A-Za-z0-9])HC-?18", RegexOptions.IgnoreCase)]
        private static partial Regex Crystal18();

        // "R_Axial_DIN0207_...", "CAP_AXIAL_P12mm", "CP_Radial_D5.0mm" - a word of its own, never inside "Coaxial".
        [GeneratedRegex(@"(?<![A-Za-z0-9])Axial(?![A-Za-z])", RegexOptions.IgnoreCase)]
        private static partial Regex Axial();

        [GeneratedRegex(@"(?<![A-Za-z0-9])Radial(?![A-Za-z])", RegexOptions.IgnoreCase)]
        private static partial Regex Radial();

        // "C_Disc_D5.0mm_...", "CAP_DISC_P100mil" - a word of its own, never inside "Discrete".
        [GeneratedRegex(@"(?<![A-Za-z0-9])Disc(?![A-Za-z])", RegexOptions.IgnoreCase)]
        private static partial Regex Disc();

        // Three legs: "EMI-Filter", "T-FILTER", and the C128's "Ferrite_Bead_vertical" (its EMI row's too).
        [GeneratedRegex(@"^(EMI-?Filter|T-?FILTER|Ferrite_Bead_vertical)", RegexOptions.IgnoreCase)]
        private static partial Regex EmiFilter();

        // Two legs, its own name saying what it is: "RESISTOR", "POWER_RESISTOR", "FERRITE_BEAD_TH_P12.7mm",
        // "INDUCTOR_P200mil", "RECTIFIER_TH_P12.7mm", "ZENER_12.7mm", "Diode".
        [GeneratedRegex(@"^((POWER_)?RESISTOR|FERRITE|INDUCTOR|RECTIFIER|ZENER|DIODE)", RegexOptions.IgnoreCase)]
        private static partial Regex LeadedByName();

        // "R_0805_2012Metric", "R_2512_6332Metric", "R_2010", "C_SMD_...", "R_MELF_MMB-0207" - no leads to speak
        // of: KiCad's "...Metric" and every chip size by its name (code review, 2026-11-04: R_2512 and R_MELF
        // read as "Axial"). By the list, as a through-hole part can hold four digits too ("DIN0207").
        [GeneratedRegex(@"(SMD|MELF|[0-9]Metric|(?<![0-9])(0201|0402|0603|0612|0805|0815|1008|1020|1206|1210|1218|1806|1812|2010|2220|2225|2512|2816|2920|4020)(?![0-9]))", RegexOptions.IgnoreCase)]
        private static partial Regex SurfaceMounted();

        // KiCad's own "Resistor_SMD", "Package_TO_SOT_SMD" - a word of its own, in the library's name.
        [GeneratedRegex(@"(?<![A-Za-z])SMD(?![A-Za-z])", RegexOptions.IgnoreCase)]
        private static partial Regex SurfaceMountLibrary();
    }
}
