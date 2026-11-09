using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH PICTURE OF THE PART A COMPONENT GETS, IF ANY (owner request, 2026-11-09: "As you already draw a resistor -
    // could you also draw an axial and radial electrolytic capacitor in same style... and a disc capacitor as a pinout
    // image and mark (if applicable) its plus and minus side? And how about transistors? And diodes? Actually all of the
    // current components for all boards"). PartPictureDrawing paints it, every position PartPictureGeometry's; this decides
    // the shape, the leads in the order they are drawn with the name under each, which lead carries the polarity mark,
    // and what is printed on the body.
    //
    // THE SHAPE COMES FROM THE PACKAGE (measured 2026-11-09 on the 14 converted boards): Disc, Radial and Axial for a
    // capacitor; DO-35, DO-41 and Axial for a diode; TO-92 and TO-220 whatever the category (transistors and regulators);
    // HC-49 and HC-18 crystals; the EMI filter; a resistor pack in SIP; a pin header; and an Axial part of any other
    // category is a ferrite bead or a coil by its names. A capacitor with NO package is drawn only when its friendly name
    // says the shape ("Electrolytic Radial", "Ceramic Axial") - the 540 "Ceramic" ones with none may be a disc, a dipped
    // or an axial part, and a wrong picture is worse than none (as the KiCad import leaves "CAP_P5.08mm" without one).
    // DIPs are PinoutDrawing's, a single resistor ResistorColorCode's and an axial ceramic capacitor whose value its color
    // code says CapacitorColorCode's: none is drawn here.
    //
    // ELECTROLYTIC OR NOT, by the FRIENDLY NAME and the value, never the description (a description on the C128DCR says
    // "not 1uF electrolytic listed on parts list" of a ceramic one): "Electrolytic" (also "elyt", "elko", "elco") is
    // polarised, "Ceramic", "Film" and the other dielectrics are not; a trimmer, a tantalum and a bipolar electrolytic are
    // not drawn. Every capacitor in the data says one; with no word, 10 uF or more is electrolytic, and anything less is
    // drawn without polarity - in doubt the plain look, never a wrong minus side (1 uF and 2.2 uF ceramic discs are in
    // the data).
    //
    // POLARITY FROM THE PINS: a pin named "+", "-", "Positive", "Negative", ... (a capacitor) or "K", "Katode", "Cathode",
    // "A", "Anode" (a diode) says which lead is which; with no such names KiCad's convention holds - a polarised
    // capacitor's pin 1 is +, a diode's pin 1 is the cathode. Checked against the board files' KiCad data: 20 of 21
    // polarised capacitors with a ground pin have it on pin 2, every diode clamping to ground or +5V has its cathode on
    // pin 1, and the data's own pin names say 1=K, 2=A. Names that contradict each other draw nothing.
    //
    // A PIN THE PICTURE HAS NO LEAD FOR (pin 3 of a disc capacitor) draws nothing: the data says something the picture
    // cannot show. Pins named rather than numbered ("B", "C", "E") cannot be placed, so their leads carry numbers only.
    // ###########################################################################################
    public static class PartPicture
    {
        public enum Shape
        {
            RadialElectrolytic,
            AxialElectrolytic,
            DiscCapacitor,
            RadialCapacitor,
            AxialCapacitor,
            GlassDiode,
            EpoxyDiode,
            TO92,
            TO220,
            Crystal,
            FerriteBead,
            Coil,
            EmiFilter,
            ResistorNetwork,
            PinHeader,
        }

        // One lead, as drawn left to right: its pin number on the case and the name written under it ("" for none).
        public sealed record Lead(int Number, string Name);

        // ###########################################################################################
        // What is drawn. MarkedLead is the index (in Leads) of the lead the part's mark is about - an electrolytic's MINUS
        // lead (a radial one's stripe stands by it and its leg is the shorter; an axial one's groove is at the OTHER end,
        // the plus end), a diode's CATHODE band, a resistor pack's pin-1 dot - null for none. Marking is what is printed on
        // the body, one line per "\n"; Caption a line under the picture saying how it is seen; Legend what the letters under
        // the leads stand for - "E" "Emitter" - in the leads' order, none when they need no explaining.
        // ###########################################################################################
        public sealed record Picture(
            Shape Shape,
            IReadOnlyList<Lead> Leads,
            int? MarkedLead,
            string Marking,
            string Caption,
            IReadOnlyList<LegendLine>? Legend = null)
        {
            public IReadOnlyList<LegendLine> LegendLines => this.Legend ?? [];
        }

        // One letter under a lead, and what it stands for.
        public sealed record LegendLine(string Letter, string Meaning);

        // ###########################################################################################
        // WHAT A TRANSISTOR'S LETTERS STAND FOR (owner request, 2026-11-09: "do have some addon info/explanation, where you
        // describe what is "E", "C" and "B""): a bipolar one's E, C, B and a field-effect one's G, D, S - a set only when
        // every lead named is one of its letters, so a regulator's "IN", "GND", "OUT" or a mix of words explains nothing.
        // ###########################################################################################
        private static readonly IReadOnlyList<IReadOnlyDictionary<string, string>> LetterSets =
        [
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["E"] = "Emitter", ["C"] = "Collector", ["B"] = "Base" },
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["G"] = "Gate", ["D"] = "Drain", ["S"] = "Source" },
        ];

        internal static IReadOnlyList<LegendLine> LegendOf(IReadOnlyList<Lead> leads)
        {
            ArgumentNullException.ThrowIfNull(leads);

            List<string> named = leads.Select(lead => lead.Name.Trim()).Where(name => name.Length > 0).ToList();

            if (named.Count == 0)
            {
                return [];
            }

            IReadOnlyDictionary<string, string>? set = LetterSets.FirstOrDefault(letters => named.All(letters.ContainsKey));

            return set is null
                ? []
                : named.Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(letter => new LegendLine(letter.ToUpperInvariant(), set[letter]))
                    .ToList();
        }

        // What the leaded packages are called (ComponentVocabulary.Packages).
        private const string Disc = "disc";
        private const string Radial = "radial";
        private const string Axial = "axial";

        private const RegexOptions Words = RegexOptions.CultureInvariant | RegexOptions.IgnoreCase;

        private static readonly Regex ElectrolyticWords = new(@"electrolyt|\belyt|\belko|\belco", Words);
        // The dielectrics that are not polarised: ceramic, and the others - one list, so the axial ceramic's rule
        // (IsAxialCeramicCapacitor) and the not-polarised one never disagree (code review, 2026-11-09).
        private const string Ceramic = "ceramic";
        private const string OtherDielectrics = @"film|polyester|mylar|polystyrene|polypropylene|\bmica\b|styroflex|\bmkt\b|\bmkp\b";

        private static readonly Regex NotPolarisedWords = new(Ceramic + "|" + OtherDielectrics, Words);
        private static readonly Regex NotDrawnCapacitorWords = new(@"trimmer|variable|adjustable|tantal|bipolar|non-?polar|\bNP\b|\bBP\b", Words);
        private static readonly Regex CeramicWords = new(Ceramic, Words);
        private static readonly Regex OtherDielectricWords = new(OtherDielectrics, Words);
        private static readonly Regex NotDrawnDiodeWords = new(@"bridge|array|\bLED\b|light.emitting|photo", Words);
        private static readonly Regex RectifierDiodes = new(@"\b1N400[1-7]\b|\b1N5[348]\d\d\b|\bBY\d|\bUF400\d|rectifier", Words);
        private static readonly Regex FerriteWords = new(@"ferrite|\bbead", Words);
        private static readonly Regex CoilWords = new(@"coil|inductor|choke", Words);

        private static readonly Regex PlusName = new(@"^(\+|\+ve|pos|positive|plus|anode)$", Words);
        private static readonly Regex MinusName = new(@"^(-|−|–|-ve|neg|negative|minus|cathode)$", Words);
        private static readonly Regex CathodeName = new(@"^(k|ka|kat|katode|kathode|cath|cathode)$", Words);
        private static readonly Regex AnodeName = new(@"^(a|an|anod|anode)$", Words);
        private static readonly Regex GroundName = new(@"^(gnd|ground|earth|0v|vss)$", Words);

        // A capacitance as typed - "100nF", "0.1 uF/25V", "10 µF/16V", "4u7", "470pF 50V 10%" - its unit in the number's
        // place too, as KiCad writes it.
        private static readonly Regex CapacitancePattern = new(
            @"^\s*(?<whole>\d*)(?:[.,](?<fraction>\d+))?\s*(?<unit>[pnuµμ])(?<after>\d+)?\s*F?(?![a-z])",
            Words);

        private static readonly Regex VoltagePattern = new(@"(?<![\d.,])(?<volts>\d+(?:[.,]\d+)?)\s*V(?![a-z])", Words);

        private static readonly Regex FrequencyPattern = new(@"(?<![\d.,])(?<number>\d+(?:[.,]\d+)?)\s*(?<unit>[mk])hz\b", Words);

        // ###########################################################################################
        // The picture of `entry` with its `pins` (of the region shown) - null when CRT draws none of it (see the header).
        // `nets`: each pin's net on this board by its number, from its KiCad data - what tells a regulator's legs apart where
        // the pins have no names (RegulatorOf); none without.
        // ###########################################################################################
        public static Picture? Of(ComponentEntry entry, IReadOnlyList<ComponentPinEntry> pins, IReadOnlyDictionary<string, string>? nets = null)
        {
            ArgumentNullException.ThrowIfNull(entry);
            ArgumentNullException.ThrowIfNull(pins);

            string category = (entry.Category ?? string.Empty).Trim();
            string package = (entry.Package ?? string.Empty).Trim().ToLowerInvariant();
            string names = $"{entry.FriendlyName} {entry.TechnicalNameOrValue}";
            var numbered = PinNames.From(pins);

            switch (package)
            {
                case "to-92":
                    return Explained(Fixed(Shape.TO92, 3, numbered, PartText(entry), "Flat side facing you, legs down"));
                case "to-220" when PositiveRegulator.Match((entry.TechnicalNameOrValue ?? string.Empty).Trim()) is { Success: true } regulator:
                    return RegulatorOf(entry, numbered, nets, int.Parse(regulator.Groups["volts"].Value, CultureInfo.InvariantCulture));
                case "to-220":
                    return Explained(Fixed(Shape.TO220, 3, numbered, PartText(entry), "Printed side facing you, legs down"));
                case "hc-49":
                case "hc-18":
                    return CrystalOf(numbered, entry);
                case "emi filter":
                    return EmiFilterOf(numbered, entry);
                case "pin header":
                    return numbered.Highest is >= 1 and <= 40
                        ? new Picture(Shape.PinHeader, Numbered(numbered, numbered.Highest), null, string.Empty, string.Empty)
                        : null;
            }

            if (Is(category, "Capacitor"))
            {
                return CapacitorOf(entry, package, names, numbered);
            }

            if (Is(category, "Diode") || Is(category, "Rectifier"))
            {
                return DiodeOf(package, names, numbered);
            }

            if (Is(category, "Resistor"))
            {
                // A pack is printed with its value, pin 1 at the dot; a single resistor is ResistorColorCode's.
                return package == "sip" && numbered.Highest is >= 2 and <= 20
                    ? new Picture(Shape.ResistorNetwork, Numbered(numbered, numbered.Highest), 0, FirstWord(entry.TechnicalNameOrValue), "Printed side facing you, pin 1 at the dot")
                    : null;
            }

            if (package == Axial)
            {
                Shape? shape = FerriteWords.IsMatch(names) ? Shape.FerriteBead
                    : CoilWords.IsMatch(names) ? Shape.Coil
                    : LabelStarts(entry, "FB") ? Shape.FerriteBead
                    : LabelStarts(entry, "L") ? Shape.Coil
                    : null;

                return shape is Shape found ? Fixed(found, 2, numbered, string.Empty, string.Empty) : null;
            }

            return null;
        }

        // ###########################################################################################
        // A capacitor: its shape by the package, or by the friendly name when it has none; electrolytic or not by its
        // words and value (see the header).
        // ###########################################################################################
        private static Picture? CapacitorOf(ComponentEntry entry, string package, string names, PinNames numbered)
        {
            // An axial ceramic one its color code says is drawn with its bands instead, as a resistor is (CapacitorColorCode,
            // shown by ResistorPicture).
            if (CapacitorColorCode.BandsOf(entry) is not null)
            {
                return null;
            }

            string shape = package.Length > 0 ? package : ShapeWordIn(entry.FriendlyName);
            decimal? picofarads = Picofarads(entry.TechnicalNameOrValue);
            string volts = Voltage(entry.TechnicalNameOrValue);

            if (shape is not (Disc or Radial or Axial) || NotDrawnCapacitorWords.IsMatch(names))
            {
                return null;
            }

            bool electrolytic = ElectrolyticWords.IsMatch(names);
            bool notPolarised = NotPolarisedWords.IsMatch(names);

            if (electrolytic && notPolarised)
            {
                return null;
            }

            if (!electrolytic && !notPolarised)
            {
                electrolytic = picofarads >= 10_000_000m;
            }

            if (shape == Disc)
            {
                return electrolytic ? null : Fixed(Shape.DiscCapacitor, 2, numbered, CodeMarking(picofarads, volts), string.Empty);
            }

            if (!electrolytic)
            {
                return Fixed(shape == Radial ? Shape.RadialCapacitor : Shape.AxialCapacitor, 2, numbered, CodeMarking(picofarads, volts), string.Empty);
            }

            string marking = picofarads is decimal value ? MicrofaradText(value) + (volts.Length > 0 ? "\n" + volts : string.Empty) : string.Empty;

            // The minus lead is the one marked; "+" and "-" under the leads where nothing else is typed.
            return Polarised(
                shape == Radial ? Shape.RadialElectrolytic : Shape.AxialElectrolytic,
                numbered,
                MinusName,
                PlusName,
                conventionalMarked: 2,
                markedRole: "-",
                otherRole: "+",
                marking);
        }

        // ###########################################################################################
        // A diode: glass DO-35 with a black band, black DO-41 with a grey one; an Axial one by its family - a rectifier
        // (1N400x, 1N54xx, ...) in black epoxy, every other (1N4148, 1N914, the zeners) in glass. The band is the cathode's.
        // ###########################################################################################
        private static Picture? DiodeOf(string package, string names, PinNames numbered)
        {
            if (NotDrawnDiodeWords.IsMatch(names))
            {
                return null;
            }

            Shape? shape = package switch
            {
                "do-35" => Shape.GlassDiode,
                "do-41" => Shape.EpoxyDiode,
                Axial => RectifierDiodes.IsMatch(names) ? Shape.EpoxyDiode : Shape.GlassDiode,
                _ => null,
            };

            // Named in full, in English (owner request, 2026-11-09: "Is that "Anode" and "Katode" ... Isn't this "Cathode"
            // instead then? ... do write the full name of it instead of a single letter").
            return shape is Shape found
                ? Polarised(found, numbered, CathodeName, AnodeName, conventionalMarked: 1, markedRole: "Cathode", otherRole: "Anode", marking: string.Empty, namesAsRoles: true)
                : null;
        }

        // ###########################################################################################
        // A 78xx POSITIVE REGULATOR IN TO-220 (owner question, 2026-11-09: "is it always implied that 1=Input, 2=Ground,
        // 3=Output? If so, then add this as info also, and there is then also a 5V variant of the same"): every maker's
        // 7805, 7812, ... stands, printed side facing you, INPUT - GROUND - OUTPUT from left to right, the tab on ground -
        // the datasheet's pins 1, 2, 3. So it is drawn by ROLE, each leg under its name ("IN", "GND", "OUT"), explained under
        // it with the voltage it gives, and numbered as THIS BOARD numbers it - which is not always the datasheet's: the C64
        // 250407's and KU-14194HB's KiCad footprint numbers VR1 and VR2 1 = input, 2 = output (the right leg), 3 = ground
        // (the middle leg and the tab), measured 2026-11-09; the C128 310378's U59 is 1, 2, 3.
        //
        // Which pin is which: the pins' own names when they say all three; else the board's KiCad nets - ground's net by its
        // name, the output's naming the voltage ("+12V", "CAN+5V" for a 7805), the input the third; else, with neither, the
        // datasheet's 1, 2, 3. Names or nets that do not settle it - two grounds, no net naming the voltage - draw nothing:
        // a leg named wrong is worse than none. A 78Lxx is another case (a TO-92 the other way round), so never matched.
        // ###########################################################################################
        private static readonly Regex PositiveRegulator = new(
            @"^(?:LM|UA|µA|μA|MC|L|KA|AN|NJM|TA)?78(?:M|S|T)?(?<volts>05|06|08|09|10|12|15|18|24)(?:[A-Z]{1,4}\d?)?$",
            Words);

        private static readonly Regex InputName = new(@"^(in|input|vin|v\s*in|vi)$", Words);
        private static readonly Regex OutputName = new(@"^(out|output|vout|v\s*out|vo)$", Words);
        private static readonly Regex RegulatorGroundName = new(@"^(gnd|ground|earth|com|common|0v)$", Words);
        private static readonly Regex GroundNet = new(@"^/?(gnd|ground|earth|0v|vss|[ad]gnd)$", Words);

        private static Picture? RegulatorOf(ComponentEntry entry, PinNames numbered, IReadOnlyDictionary<string, string>? nets, int volts)
        {
            if (numbered.Highest > 3)
            {
                return null;
            }

            (int Input, int Ground, int Output)? roles = RegulatorRoles(numbered, nets, volts);

            if (roles is not { } legs)
            {
                return null;
            }

            string voltage = string.Create(CultureInfo.InvariantCulture, $"+{volts} V");

            return new Picture(
                Shape.TO220,
                [new Lead(legs.Input, "IN"), new Lead(legs.Ground, "GND"), new Lead(legs.Output, "OUT")],
                null,
                PartText(entry),
                "Printed side facing you, legs down",
                [new LegendLine("IN", "Input"), new LegendLine("GND", "Ground"), new LegendLine("OUT", $"Output, {voltage}")]);
        }

        // ###########################################################################################
        // Which of the board's pins is the input, the ground and the output (RegulatorOf's header) - null when that is not
        // settled.
        // ###########################################################################################
        private static (int Input, int Ground, int Output)? RegulatorRoles(PinNames numbered, IReadOnlyDictionary<string, string>? nets, int volts)
        {
            int[] pins = [1, 2, 3];
            List<string> typed = pins.Select(pin => numbered.LeadOf(pin).Name).ToList();

            if (typed.Any(name => name.Length > 0))
            {
                // Named: all three, each once, or nothing is drawn.
                int? Named(Regex role) => pins.Where(pin => role.IsMatch(typed[pin - 1])).ToList() is [int only] ? only : null;

                return Named(InputName) is int input && Named(RegulatorGroundName) is int ground && Named(OutputName) is int output &&
                       new[] { input, ground, output }.Distinct().Count() == 3
                    ? (input, ground, output)
                    : null;
            }

            if (nets is null || nets.Count == 0)
            {
                return (1, 2, 3);
            }

            string NetOf(int pin) => nets.TryGetValue(pin.ToString(CultureInfo.InvariantCulture), out string? net) ? net.Trim() : string.Empty;

            Regex voltageNet = VoltageNetOf(volts);
            List<int> grounds = pins.Where(pin => GroundNet.IsMatch(NetOf(pin))).ToList();
            List<int> outputs = pins.Where(pin => !grounds.Contains(pin) && voltageNet.IsMatch(NetOf(pin)) &&
                                                  !NetOf(pin).Contains("UNREG", StringComparison.OrdinalIgnoreCase)).ToList();

            if (grounds is not [int groundPin] || outputs is not [int outputPin])
            {
                return null;
            }

            return (pins.Single(pin => pin != groundPin && pin != outputPin), groundPin, outputPin);
        }

        // ###########################################################################################
        // The net a 78xx of `volts` puts out on - "+5V", "5V_REG", never "15V" or "1.5V". Made once per voltage (code review,
        // 2026-11-09: a new Regex at every drawing, which the pane makes at every pause in typing and every KiCad answer).
        // ###########################################################################################
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<int, Regex> VoltageNets = new();

        private static Regex VoltageNetOf(int volts) =>
            VoltageNets.GetOrAdd(volts, each => new Regex(string.Create(CultureInfo.InvariantCulture, $@"(?<![\d.]){each}V"), Words));

        // A crystal can: two leads, or three with the case's own in the middle (pin 3 - the KiCad footprints' third pad).
        private static Picture? CrystalOf(PinNames numbered, ComponentEntry entry)
        {
            string marking = FrequencyText(entry);

            if (numbered.Highest <= 2)
            {
                return Fixed(Shape.Crystal, 2, numbered, marking, string.Empty);
            }

            return numbered.Highest == 3
                ? new Picture(Shape.Crystal, [numbered.LeadOf(1), numbered.LeadOf(3), numbered.LeadOf(2)], null, marking, string.Empty)
                : null;
        }

        // ###########################################################################################
        // An EMI filter's three leads stand 1, 2, 3 - unless one is named for ground, which is always the middle one: the
        // Open128's ferrite-bead filters have their ground on pin 3 between 1 and 2, its EMI filters on pin 2.
        // ###########################################################################################
        private static Picture? EmiFilterOf(PinNames numbered, ComponentEntry entry)
        {
            if (numbered.Highest > 3)
            {
                return null;
            }

            List<Lead> leads = [numbered.LeadOf(1), numbered.LeadOf(2), numbered.LeadOf(3)];
            List<Lead> grounds = leads.Where(lead => GroundName.IsMatch(lead.Name)).ToList();

            if (grounds.Count == 1)
            {
                leads.Remove(grounds[0]);
                leads.Insert(1, grounds[0]);
            }

            return new Picture(Shape.EmiFilter, leads, null, CodeMarking(Picofarads(entry.TechnicalNameOrValue), string.Empty), string.Empty);
        }

        // ###########################################################################################
        // A two-lead part with a polarity: the lead whose name says `marked` (an electrolytic's minus, a diode's
        // cathode), or the other's name says `other`, or neither - KiCad's convention, `conventionalMarked`. Names that
        // say both of one lead, or the same of both, draw nothing. A lead with no name is written as its role; with
        // `namesAsRoles` (a diode's), one whose name says it however it was typed ("K", "KATODE") is too - "Cathode",
        // "Anode", in full and in English (owner request, 2026-11-09). Any other name as typed.
        // ###########################################################################################
        private static Picture? Polarised(
            Shape shape,
            PinNames numbered,
            Regex marked,
            Regex other,
            int conventionalMarked,
            string markedRole,
            string otherRole,
            string marking,
            bool namesAsRoles = false)
        {
            if (numbered.Highest > 2)
            {
                return null;
            }

            Lead first = numbered.LeadOf(1);
            Lead second = numbered.LeadOf(2);

            bool firstMarked = marked.IsMatch(first.Name), secondMarked = marked.IsMatch(second.Name);
            bool firstOther = other.IsMatch(first.Name), secondOther = other.IsMatch(second.Name);

            int? markedNumber =
                firstMarked && !secondMarked && !firstOther ? 1
                : secondMarked && !firstMarked && !secondOther ? 2
                : !firstMarked && !secondMarked && firstOther && !secondOther ? 2
                : !firstMarked && !secondMarked && secondOther && !firstOther ? 1
                : !firstMarked && !secondMarked && !firstOther && !secondOther ? conventionalMarked
                : null;

            if (markedNumber is not int markedPin)
            {
                return null;
            }

            Lead Named(Lead lead) =>
                lead.Name.Length == 0 || (namesAsRoles && (marked.IsMatch(lead.Name) || other.IsMatch(lead.Name)))
                    ? lead with { Name = lead.Number == markedPin ? markedRole : otherRole }
                    : lead;

            return new Picture(shape, [Named(first), Named(second)], markedPin - 1, marking, string.Empty);
        }

        // The picture with what its leads' letters stand for, when they need explaining (LegendOf).
        private static Picture? Explained(Picture? picture) =>
            picture is null ? null : picture with { Legend = PartPicture.LegendOf(picture.Leads) };

        // A part with `count` leads numbered 1 to `count`; a pin past the last draws nothing.
        private static Picture? Fixed(Shape shape, int count, PinNames numbered, string marking, string caption) =>
            numbered.Highest > count ? null : new Picture(shape, Numbered(numbered, count), null, marking, caption);

        private static List<Lead> Numbered(PinNames numbered, int count) =>
            Enumerable.Range(1, count).Select(numbered.LeadOf).ToList();

        // ###########################################################################################
        // THE PINS BY NUMBER: the first name typed for each, trimmed - a name that only repeats its number ("1" for pin
        // 1, as 49 of the data's EMI filters and beads have it) is none - and the highest number named.
        // ###########################################################################################
        private sealed class PinNames
        {
            private readonly Dictionary<int, string> thisNames = new();

            public int Highest { get; private set; }

            public static PinNames From(IReadOnlyList<ComponentPinEntry> pins)
            {
                var result = new PinNames();

                foreach (ComponentPinEntry pin in pins)
                {
                    if (ComponentVocabulary.PinNumberOf(pin.Pin) is not int number || number < 1)
                    {
                        continue;
                    }

                    string name = (pin.Name ?? string.Empty).Trim();

                    if (name == number.ToString(CultureInfo.InvariantCulture))
                    {
                        name = string.Empty;
                    }

                    result.thisNames.TryAdd(number, name);
                    result.Highest = Math.Max(result.Highest, number);
                }

                return result;
            }

            public Lead LeadOf(int number) => new(number, this.thisNames.GetValueOrDefault(number, string.Empty));
        }

        // ###########################################################################################
        // A capacitance in picofarads - null when the value says none. Internal for the tests.
        // ###########################################################################################
        internal static decimal? Picofarads(string? value)
        {
            string text = value ?? string.Empty;
            Match match = CapacitancePattern.Match(text);

            // "10,000pF" may be ten thousand or ten (code review, 2026-11-09 - read as 10 pF, its bands and printed code a
            // thousand times out): no value, as a resistor's "1,000" (ResistorColorCode.CommaMayBeThousands).
            if (!match.Success || (match.Groups["whole"].Length == 0 && !match.Groups["fraction"].Success) ||
                (match.Groups["after"].Success && match.Groups["fraction"].Success) ||
                ResistorColorCode.CommaMayBeThousands(text, match.Groups["fraction"]))
            {
                return null;
            }

            string digits = (match.Groups["whole"].Length > 0 ? match.Groups["whole"].Value : "0") +
                            (match.Groups["fraction"].Success ? "." + match.Groups["fraction"].Value
                                : match.Groups["after"].Success ? "." + match.Groups["after"].Value
                                : string.Empty);

            if (digits.Length > 12 || !decimal.TryParse(digits, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal number) || number <= 0)
            {
                return null;
            }

            return match.Groups["unit"].Value.ToLowerInvariant() switch
            {
                "p" => number,
                "n" => number * 1_000m,
                _ => number * 1_000_000m,
            };
        }

        // The voltage typed after the value - "16V" - or "".
        internal static string Voltage(string? value)
        {
            Match match = VoltagePattern.Match(value ?? string.Empty);

            return match.Success ? match.Groups["volts"].Value.Replace(',', '.') + "V" : string.Empty;
        }

        // ###########################################################################################
        // A capacitor's three-figure code, as a ceramic one is printed: two digits and the number of zeros after them, in
        // picofarads - 100 nF "104", 470 pF "471", 1.8 nF "182" - and a value under 100 pF as its number ("47"). "" for one
        // two digits cannot say.
        // ###########################################################################################
        internal static string CapacitorCode(decimal picofarads)
        {
            if (picofarads <= 0)
            {
                return string.Empty;
            }

            if (picofarads < 100)
            {
                return picofarads.ToString("0.#", CultureInfo.InvariantCulture);
            }

            for (int zeros = 1; zeros <= 9; zeros++)
            {
                decimal digits = picofarads / Pow10(zeros);

                if (digits >= 10 && digits < 100)
                {
                    return digits == decimal.Truncate(digits) ? $"{(int)digits}{zeros}" : string.Empty;
                }
            }

            return string.Empty;
        }

        private static decimal Pow10(int exponent)
        {
            decimal result = 1;

            for (int i = 0; i < exponent; i++)
            {
                result *= 10;
            }

            return result;
        }

        // The code, and the voltage under it when typed.
        private static string CodeMarking(decimal? picofarads, string volts)
        {
            string code = picofarads is decimal value ? CapacitorCode(value) : string.Empty;

            return code.Length == 0 ? string.Empty : volts.Length > 0 ? code + "\n" + volts : code;
        }

        // An electrolytic's capacitance as it is printed on the can: "10µF", "4.7µF", "0.1µF".
        internal static string MicrofaradText(decimal picofarads) =>
            (picofarads / 1_000_000m).ToString("0.###", CultureInfo.InvariantCulture) + "µF";

        // A crystal's frequency, from its value, friendly name or description ("17.73447 MHz (PAL)"): "17.73447\nMHz".
        private static string FrequencyText(ComponentEntry entry)
        {
            foreach (string text in new[] { entry.TechnicalNameOrValue, entry.FriendlyName, entry.Description })
            {
                Match match = FrequencyPattern.Match(text ?? string.Empty);

                if (match.Success)
                {
                    bool mega = match.Groups["unit"].Value.Equals("m", StringComparison.OrdinalIgnoreCase);

                    return match.Groups["number"].Value.Replace(',', '.') + "\n" + (mega ? "MHz" : "kHz");
                }
            }

            return string.Empty;
        }

        // What a transistor or regulator has printed on it: its type ("2SC1815", "7812"), else its part number - a short
        // word only.
        private static string PartText(ComponentEntry entry)
        {
            foreach (string text in new[] { entry.TechnicalNameOrValue, entry.PartNumber })
            {
                string trimmed = (text ?? string.Empty).Trim();

                if (trimmed.Length is > 0 and <= 10 && !trimmed.Any(char.IsWhiteSpace))
                {
                    return trimmed;
                }
            }

            return string.Empty;
        }

        // The first word of a value ("3K3Ω, 6pin" is "3K3Ω"), when short enough to print.
        private static string FirstWord(string? value)
        {
            string word = (value ?? string.Empty).Trim().Split([' ', ',', ';', '/'], StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;

            return word.Length <= 10 ? word : string.Empty;
        }

        // ###########################################################################################
        // Whether `entry` is an axial ceramic capacitor - its package "Axial", or with none its friendly name saying so as
        // CapacitorOf reads it; "Ceramic" in its names and no other dielectric, no electrolytic, nothing CapacitorOf draws
        // nothing of (a trimmer). The kind CapacitorColorCode reads bands for (owner decision, 2026-11-09: "Every axial
        // ceramic").
        // ###########################################################################################
        internal static bool IsAxialCeramicCapacitor(ComponentEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            if (!Is((entry.Category ?? string.Empty).Trim(), "Capacitor"))
            {
                return false;
            }

            string package = (entry.Package ?? string.Empty).Trim().ToLowerInvariant();
            string shape = package.Length > 0 ? package : ShapeWordIn(entry.FriendlyName);
            string names = $"{entry.FriendlyName} {entry.TechnicalNameOrValue}";

            return shape == Axial &&
                   CeramicWords.IsMatch(names) &&
                   !OtherDielectricWords.IsMatch(names) &&
                   !ElectrolyticWords.IsMatch(names) &&
                   !NotDrawnCapacitorWords.IsMatch(names);
        }

        // The shape a friendly name says, when it says exactly one - "Electrolytic Radial" - else "".
        private static string ShapeWordIn(string? friendlyName)
        {
            string text = friendlyName ?? string.Empty;
            var shapes = new List<string>();

            if (Regex.IsMatch(text, @"\bradial\b", Words))
            {
                shapes.Add(Radial);
            }

            if (Regex.IsMatch(text, @"\baxial\b", Words))
            {
                shapes.Add(Axial);
            }

            if (Regex.IsMatch(text, @"\bdis[ck]\b", Words))
            {
                shapes.Add(Disc);
            }

            return shapes.Count == 1 ? shapes[0] : string.Empty;
        }

        private static bool Is(string category, string wanted) => string.Equals(category, wanted, StringComparison.OrdinalIgnoreCase);

        // Whether the label starts with `prefix` and a digit - "FB12", "L3".
        private static bool LabelStarts(ComponentEntry entry, string prefix)
        {
            string label = (entry.BoardLabel ?? string.Empty).Trim();

            return label.Length > prefix.Length &&
                   label.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                   char.IsAsciiDigit(label[prefix.Length]);
        }
    }
}
