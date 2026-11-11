using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // The fields of a component's own row the component pane edits one at a time.
    public enum ComponentField
    {
        FriendlyName,
        TechnicalNameOrValue,
        PartNumber,
        Category,
        Package,
        Description,
        Purpose,
    }

    // ###########################################################################################
    // WHAT THE COMPONENT PANE CHANGES IN A DRAFT (CRT 4.0 step 2, 2026-10-25 - the plan: "A
    // component opens in a pane, never in a maximized window ... Basics: a one-line short description
    // with a counter ... Pinout ... Replacements").
    //
    // The pane edits a component as it is SHOWN: the region CRT shows, by the component list's own
    // rule (ComponentListBuilder - a row with a blank region is in every region). Every edit is one
    // field, or one section's rows, and arrives here as a transform of the draft's board - applied to
    // the draft held in memory as one step of its history (DraftSession), so it is one Undo away and
    // the table shows it at once.
    //
    //   - A FIELD changes the component's row in the region shown, and nothing else.
    //   - A SECTION (pins, replacement parts, links) replaces the component's rows of that section in
    //     the region shown, each in its own place (ReplaceInScope) - the other region's rows, and every
    //     other component's, are carried across untouched (ComponentBoardWriter's lesson: "Before
    //     assuming a section can be replaced wholesale, check how it was loaded").
    //   - An edit that changes nothing returns THE SAME board, which DraftSession takes as no step.
    //
    // PURE: takes a board, returns a board, touches no files.
    // ###########################################################################################
    public static class ComponentPaneEdits
    {
        // What a field is called on the pane, and in the draft's history ("Other technical details of [U8]" - "Short
        // description" until 2026-11-09, owner request).
        public static string NameOf(ComponentField field) => field switch
        {
            ComponentField.FriendlyName => "Friendly name",
            ComponentField.TechnicalNameOrValue => "Technical name or value",
            ComponentField.PartNumber => "Part number",
            ComponentField.Category => "Category",
            ComponentField.Package => "Package",
            ComponentField.Description => BoardWorkbookSchema.DescriptionHeading,
            ComponentField.Purpose => BoardWorkbookSchema.ColPurpose,
            _ => throw new ArgumentOutOfRangeException(nameof(field), field, null),
        };

        // ###########################################################################################
        // The component's row in `region` - the one the component list shows there - or, when none
        // matches the region, its first row: what the component editor's window has always opened on.
        // Null when the board has no component of that label.
        // ###########################################################################################
        public static ComponentEntry? RowShown(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return null;
            }

            List<ComponentEntry> rows = board.Components.Where(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted)).ToList();

            return rows.FirstOrDefault(row => ComponentPaneEdits.InRegion(row.Region, region)) ?? rows.FirstOrDefault();
        }

        public static string ValueOf(ComponentEntry row, ComponentField field)
        {
            ArgumentNullException.ThrowIfNull(row);

            return field switch
            {
                ComponentField.FriendlyName => row.FriendlyName,
                ComponentField.TechnicalNameOrValue => row.TechnicalNameOrValue,
                ComponentField.PartNumber => row.PartNumber,
                ComponentField.Category => row.Category,
                ComponentField.Package => row.Package,
                ComponentField.Description => row.Description,
                ComponentField.Purpose => row.Purpose,
                _ => throw new ArgumentOutOfRangeException(nameof(field), field, null),
            } ?? string.Empty;
        }

        // ###########################################################################################
        // A field's value as it is kept: trimmed, and on one line - but "Purpose or note", a note with formatting, keeps its
        // lines ("\n" alone) and is trimmed as such a note is (RichNote.TrimMarkup: a first line's indent says a bullet).
        // ###########################################################################################
        public static string Normalised(ComponentField field, string? value)
        {
            string text = value ?? string.Empty;

            if (field == ComponentField.Purpose)
            {
                return RichNote.TrimMarkup(text.Replace("\r\n", "\n").Replace('\r', '\n'));
            }

            return text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
        }

        // A note of several lines as it is kept: trimmed, each line ended by "\n" alone, whatever the
        // clipboard brought. A replacement part's - plain text, typed in a plain box.
        public static string NormalisedNote(string? value) =>
            (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Trim();

        // ###########################################################################################
        // The board with one field of the component's row in `region` set - the same board when the
        // field already reads so, or the board has no such component.
        // ###########################################################################################
        public static BoardData WithField(BoardData board, string? label, string? region, ComponentField field, string? value)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (ComponentPaneEdits.RowShown(board, label, region) is not ComponentEntry row)
            {
                return board;
            }

            string wanted = ComponentPaneEdits.Normalised(field, value);

            // Compared as a field is kept (code review, 2026-11-02): C16 251789's R10 holds a line break, and
            // its box handed back as it was shown read as a change nobody typed - a step and a "Draft" mark
            // for opening the component and leaving it.
            if (string.Equals(ComponentPaneEdits.Normalised(field, ComponentPaneEdits.ValueOf(row, field)), wanted, StringComparison.Ordinal))
            {
                return board;
            }

            var changed = new ComponentEntry
            {
                BoardLabel = row.BoardLabel,
                Region = row.Region,
                FriendlyName = field == ComponentField.FriendlyName ? wanted : row.FriendlyName,
                TechnicalNameOrValue = field == ComponentField.TechnicalNameOrValue ? wanted : row.TechnicalNameOrValue,
                PartNumber = field == ComponentField.PartNumber ? wanted : row.PartNumber,
                Category = field == ComponentField.Category ? wanted : row.Category,
                Package = field == ComponentField.Package ? wanted : row.Package,
                Description = field == ComponentField.Description ? wanted : row.Description,
                Purpose = field == ComponentField.Purpose ? wanted : row.Purpose,
            };

            List<ComponentEntry> components = board.Components.Select(entry => ReferenceEquals(entry, row) ? changed : entry).ToList();

            return board.WithComponentSections(components: components);
        }

        // ---- a new component (slice 3c) -----------------------------------------------------------

        // ###########################################################################################
        // The board with a new component - its label, its region (blank: every region) and its
        // category, nothing else yet - placed where a new component goes (ComponentPlacement: into its
        // category in label order, a new category last). The same board when the label is blank or the
        // board has it already, in any region: a label names one component (ContributionPackaging
        // .ValidateNewComponent says why to the person adding it).
        // ###########################################################################################
        public static BoardData WithNewComponent(BoardData board, string? label, string? region, string? category)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = ComponentPaneEdits.Normalised(ComponentField.FriendlyName, label);

            if (wanted.Length == 0 || board.Components.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted)))
            {
                return board;
            }

            string filedCategory = ComponentPaneEdits.Normalised(ComponentField.Category, category);

            var components = board.Components.ToList();
            components.Insert(
                ComponentPlacement.InsertionIndex(components, filedCategory, wanted),
                new ComponentEntry
                {
                    BoardLabel = wanted,
                    Region = region?.Trim() ?? string.Empty,
                    Category = filedCategory,
                });

            return board.WithComponentSections(components: components);
        }

        // ---- its label (owner report, 2026-11-01: "unable to rename the component label") ----------

        // ###########################################################################################
        // The board with the component `from` under the label `to`: every row of it, every region, its
        // pictures, baselines, pins, replacement parts, files, links and rectangles. Its rectangles are
        // moved here - DraftEditRules follows a relabel only for a row whose other cells stayed, and one
        // that has twin rows or another change would have left them behind.
        //
        // The same board when `to` is blank, says the same, or names another component in any region -
        // a label names one component (ContributionPackaging.ValidateNewComponent says why to the person
        // typing). Capitals alone are a change, as what the board calls it.
        // ###########################################################################################
        public static BoardData WithLabel(BoardData board, string? from, string? to)
        {
            ArgumentNullException.ThrowIfNull(board);

            string source = from?.Trim() ?? string.Empty;
            string wanted = ComponentPaneEdits.Normalised(ComponentField.FriendlyName, to);

            if (source.Length == 0 || wanted.Length == 0 || string.Equals(source, wanted, StringComparison.Ordinal) ||
                !board.Components.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source)))
            {
                return board;
            }

            if (!ComponentPaneEdits.SameLabel(source, wanted) &&
                board.Components.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted)))
            {
                return board;
            }

            return board
                .WithComponentSections(
                    components: board.Components.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList(),
                    componentImages: board.ComponentImages.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList(),
                    componentLocalFiles: board.ComponentLocalFiles.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList(),
                    componentLinks: board.ComponentLinks.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList(),
                    componentPins: board.ComponentPins.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList(),
                    componentReplacements: board.ComponentReplacements.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList())
                .WithComponentHighlights(board.ComponentHighlights.Select(row => ComponentPaneEdits.SameLabel(row.BoardLabel, source) ? row.WithBoardLabel(wanted) : row).ToList());
        }

        // ---- told apart per region (owner decision, 2026-11-01) --------------------------------------

        // Whether the component has a row of its own per region.
        public static bool IsRegional(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.Components.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, label?.Trim() ?? string.Empty) && !string.IsNullOrWhiteSpace(row.Region));
        }

        // ###########################################################################################
        // The component told apart per region, or no longer. On: a row per region, each starting as the
        // one row it had. Off: the row `region` shows is the one row, for every region, and the others
        // go - "Undo" brings them back. One step; the same board when it already reads so.
        // ###########################################################################################
        public static BoardData WithRegionalRows(BoardData board, string? label, string? region, bool regional)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            // No region to tell it apart by, nothing is (code review, 2026-11-09): a row per region of none was no row at all,
            // and the component went.
            if (ComponentPaneEdits.RowShown(board, wanted, region) is not ComponentEntry shown ||
                ComponentPaneEdits.IsRegional(board, wanted) == regional ||
                (regional && BoardRegions.Of(board).Count == 0))
            {
                return board;
            }

            List<ComponentEntry> rows = regional
                ? BoardRegions.Of(board).Select(each => shown.WithRegion(each)).ToList()
                : [shown.WithRegion(string.Empty)];

            return board.WithComponentSections(components: ComponentPaneEdits.ReplaceInScope(
                board.Components,
                row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted),
                rows));
        }

        // Whether the component's pins differ per region - a set of pins of each's own.
        public static bool PinsAreRegional(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentPins.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, label?.Trim() ?? string.Empty) && !string.IsNullOrWhiteSpace(row.Region));
        }

        // ###########################################################################################
        // The component's pins told apart per region, or no longer. On: each region starts with a copy
        // of the pins for every region. Off: the pins `region` shows are the pins of every region, the
        // others go. One step; the same board when it already reads so, when there is nothing to tell
        // apart yet - the first pin typed afterwards is the row's own - and when `region` shows no pins:
        // made the same from nothing, every pin of the component went (code review, 2026-11-01).
        // ###########################################################################################
        public static BoardData WithRegionalPins(BoardData board, string? label, string? region, bool regional)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            if (ComponentPaneEdits.PinsAreRegional(board, wanted) == regional ||
                !board.ComponentPins.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted)) ||
                (regional && BoardRegions.Of(board).Count == 0))
            {
                return board;
            }

            List<ComponentPinEntry> shown = ComponentPaneEdits.PinsOf(board, wanted, region);

            if (!regional && shown.Count == 0)
            {
                return board;
            }

            List<ComponentPinEntry> rows = regional
                ? BoardRegions.Of(board).SelectMany(each => board.ComponentPins
                        .Where(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted))
                        .Select(row => row.WithRegion(each)))
                    .ToList()
                : shown.Select(row => row.WithRegion(string.Empty)).ToList();

            return board.WithComponentSections(componentPins: ComponentPaneEdits.ReplaceInScope(
                board.ComponentPins,
                row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted),
                rows));
        }

        // Whether the component has a highlight on any picture - what "Show on board" can show.
        public static bool IsMarkedOnAPicture(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            return wanted.Length > 0 && board.ComponentHighlights.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted));
        }

        // ---- pins -------------------------------------------------------------------------------

        // ###########################################################################################
        // The region a NEW pin of `label` is filed under, CRT showing `region`: it joins the set it stands with - the region's
        // own while the pins are told apart, else the one for every region - but a component that is in one region only: its
        // pins are that region's, the region it is worked on in. Asked of each component a pin is added to (code review,
        // 2026-11-09: with several picked the first one's was given to all of them).
        // ###########################################################################################
        public static string RegionForNewPin(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            List<ComponentEntry> rows = board.Components.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).ToList();

            // A component of one region works in that region whatever region CRT shows, its pins told apart or not (second code
            // review, 2026-11-09: picked with others and the region switched, its new pin went to a region it is not in -
            // shown nowhere, and sent so).
            if (rows.Count == 1 && !string.IsNullOrWhiteSpace(rows[0].Region))
            {
                return rows[0].Region.Trim();
            }

            return ComponentPaneEdits.PinsAreRegional(board, label) ? region?.Trim() ?? string.Empty : string.Empty;
        }

        // The component's pins in `region`, as the board lists them.
        public static List<ComponentPinEntry> PinsOf(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentPins.Where(pin => ComponentPaneEdits.InScope(pin.BoardLabel, pin.Region, label, region)).ToList();
        }

        // ###########################################################################################
        // The board with the component's pins in `region` replaced by `pins` - each filed under the
        // component's label; a pin with nothing in it is left out. A new pin carries the region of
        // the component's row (blank for a component in every region); a pin the board had keeps its
        // own. The same board when nothing changed.
        // ###########################################################################################
        public static BoardData WithPins(BoardData board, string? label, string? region, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(pins);

            List<ComponentPinEntry> before = ComponentPaneEdits.PinsOf(board, label, region);
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region, before.Select(pin => pin.BoardLabel));

            List<ComponentPinEntry> rows = pins
                .Select(pin => new ComponentPinEntry
                {
                    BoardLabel = filedLabel,
                    Region = pin.Region?.Trim() ?? string.Empty,
                    Pin = pin.Pin?.Trim() ?? string.Empty,
                    Name = pin.Name?.Trim() ?? string.Empty,
                    Direction = pin.Direction?.Trim() ?? string.Empty,
                })
                .Where(pin => pin.Pin.Length > 0 || pin.Name.Length > 0 || pin.Direction.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, ComponentPaneEdits.KeyOf))
            {
                return board;
            }

            return board.WithComponentSections(componentPins: ComponentPaneEdits.ReplaceInScope(
                board.ComponentPins,
                pin => ComponentPaneEdits.InScope(pin.BoardLabel, pin.Region, label, region),
                rows));
        }

        // ###########################################################################################
        // `pins` with every number of a `count`-pin case present - the ones missing added with
        // nothing but their number, in `region` - and all of them in number order. For "the pins of
        // the case" once a package is picked (ComponentVocabulary.PinCountOf).
        // ###########################################################################################
        public static List<ComponentPinEntry> WithEveryPinOf(IReadOnlyList<ComponentPinEntry> pins, int count, string? region)
        {
            ArgumentNullException.ThrowIfNull(pins);

            var rows = pins.ToList();
            var numbers = new HashSet<string>(rows.Select(pin => pin.Pin?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);

            for (int number = 1; number <= count; number++)
            {
                string text = number.ToString(System.Globalization.CultureInfo.InvariantCulture);

                if (numbers.Add(text))
                {
                    rows.Add(new ComponentPinEntry { Region = region?.Trim() ?? string.Empty, Pin = text });
                }
            }

            return rows.OrderBy(pin => pin.Pin, ComponentPaneEdits.PinOrder).ToList();
        }

        // ###########################################################################################
        // Pins in the order of their numbers: "2" before "10", a number before a name ("A1"), names
        // by their letters - a pin with no number last.
        // ###########################################################################################
        public static IComparer<string?> PinOrder { get; } = Comparer<string?>.Create(ComparePins);

        private static int ComparePins(string? first, string? second)
        {
            string a = first?.Trim() ?? string.Empty;
            string b = second?.Trim() ?? string.Empty;

            if (a.Length == 0 || b.Length == 0)
            {
                return (a.Length == 0).CompareTo(b.Length == 0);
            }

            bool aNumber = int.TryParse(a, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int aValue);
            bool bNumber = int.TryParse(b, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int bValue);

            if (aNumber && bNumber)
            {
                return aValue.CompareTo(bValue);
            }

            if (aNumber != bNumber)
            {
                return aNumber ? -1 : 1;
            }

            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        }

        // ---- replacement parts ------------------------------------------------------------------

        public static List<ComponentReplacementEntry> ReplacementsOf(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentReplacements.Where(row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, region)).ToList();
        }

        // The same as WithPins, for the component's replacement parts.
        public static BoardData WithReplacements(BoardData board, string? label, string? region, IReadOnlyList<ComponentReplacementEntry> replacements)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(replacements);

            List<ComponentReplacementEntry> before = ComponentPaneEdits.ReplacementsOf(board, label, region);
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region, before.Select(row => row.BoardLabel));

            List<ComponentReplacementEntry> rows = replacements
                .Select(row => new ComponentReplacementEntry
                {
                    BoardLabel = filedLabel,
                    Region = row.Region?.Trim() ?? string.Empty,
                    PartNumber = row.PartNumber?.Trim() ?? string.Empty,
                    Kind = row.Kind?.Trim() ?? string.Empty,
                    Note = ComponentPaneEdits.NormalisedNote(row.Note),
                })
                .Where(row => row.PartNumber.Length > 0 || row.Kind.Length > 0 || row.Note.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, ComponentPaneEdits.KeyOf))
            {
                return board;
            }

            return board.WithComponentSections(componentReplacements: ComponentPaneEdits.ReplaceInScope(
                board.ComponentReplacements,
                row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, region),
                rows));
        }

        // ---- links ------------------------------------------------------------------------------

        // ###########################################################################################
        // EVERY link of the component, whatever region CRT shows: its "Links" page is the component's whole
        // list, as "Board: Links" is the board's, each row saying what it applies to (owner request,
        // 2026-11-03: "Each file or link will have set a region (if any) and then the complete list can be
        // viewed and sorted"). What a region shows of them is the popup's (ComponentPopupContent.LinksOf).
        // ###########################################################################################
        public static List<ComponentLinkEntry> AllLinksOf(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentLinks.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).ToList();
        }

        // The component's whole list of links replaced by `links`, each in its own place.
        public static BoardData WithAllLinks(BoardData board, string? label, IReadOnlyList<ComponentLinkEntry> links)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(links);

            if (string.IsNullOrWhiteSpace(label))
            {
                return board;
            }

            bool inScope(ComponentLinkEntry row) => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label);

            List<ComponentLinkEntry> before = board.ComponentLinks.Where(inScope).ToList();
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region: null, before.Select(row => row.BoardLabel));

            List<ComponentLinkEntry> rows = links
                .Select(row => new ComponentLinkEntry
                {
                    BoardLabel = filedLabel,
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                })
                .Where(row => row.Name.Length > 0 || row.Url.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, row => string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Region?.Trim(), row.Name?.Trim(), row.Url?.Trim())))
            {
                return board;
            }

            return board.WithComponentSections(componentLinks: ComponentPaneEdits.ReplaceInScope(board.ComponentLinks, inScope, rows));
        }

        // ---- pictures and scope baselines (slice 3b) ---------------------------------------------

        // The component's pictures in `region`: its "Component images" rows that are no scope
        // baseline (ComponentImageKinds - the one rule, the popup's and the converter's).
        public static List<ComponentImageEntry> PicturesOf(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentImages
                .Where(row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, region) && !ComponentImageKinds.IsScopeBaseline(row))
                .ToList();
        }

        // The component's scope baselines in `region`: the rows naming a pin.
        public static List<ComponentImageEntry> BaselinesOf(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentImages
                .Where(row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, region) && ComponentImageKinds.IsScopeBaseline(row))
                .ToList();
        }

        // ###########################################################################################
        // The board with the component's pictures in `region` replaced by `pictures`, each in its own
        // place - its scope baselines and every other row left in place (one list holds both).
        // Each filed under the component's label; a row with nothing in it is left out; every value
        // trimmed, a note keeping its lines. The same board when nothing changed.
        // ###########################################################################################
        public static BoardData WithPictures(BoardData board, string? label, string? region, IReadOnlyList<ComponentImageEntry> pictures) =>
            ComponentPaneEdits.WithImages(board, label, region, pictures, baselines: false);

        // ###########################################################################################
        // The same for the component's scope baselines, its pictures left in place. A row handed over
        // that is no baseline by the rule is written as it is - read as a picture from then on: the pane
        // holds back an edit that would do that (a pin cleared), and nothing here drops a row.
        // ###########################################################################################
        public static BoardData WithBaselines(BoardData board, string? label, string? region, IReadOnlyList<ComponentImageEntry> baselines) =>
            ComponentPaneEdits.WithImages(board, label, region, baselines, baselines: true);

        private static BoardData WithImages(BoardData board, string? label, string? region, IReadOnlyList<ComponentImageEntry> images, bool baselines)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(images);

            List<ComponentImageEntry> before = baselines
                ? ComponentPaneEdits.BaselinesOf(board, label, region)
                : ComponentPaneEdits.PicturesOf(board, label, region);
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region, before.Select(row => row.BoardLabel));

            List<ComponentImageEntry> rows = images
                .Select(row => new ComponentImageEntry
                {
                    BoardLabel = filedLabel,
                    Region = row.Region?.Trim() ?? string.Empty,
                    Pin = row.Pin?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty,
                    // Markup, typed in a formatted box: its first line's indent says whether it is a bullet and how deep,
                    // so it is trimmed as notes are, never Trim() (RichNote.NormalisedMarkup; code review, 2026-11-08).
                    Note = RichNote.NormalisedMarkup(row.Note),
                    TimeDiv = row.TimeDiv?.Trim() ?? string.Empty,
                    VoltsDiv = row.VoltsDiv?.Trim() ?? string.Empty,
                    TriggerLevelVolts = row.TriggerLevelVolts?.Trim() ?? string.Empty,
                    MachineState = row.MachineState?.Trim() ?? string.Empty,
                })
                .Where(row => row.Name.Length > 0 || row.File.Length > 0 || row.Note.Length > 0 || row.Pin.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, ComponentPaneEdits.KeyOf))
            {
                return board;
            }

            return board.WithComponentSections(componentImages: ComponentPaneEdits.ReplaceInScope(
                board.ComponentImages,
                row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, region) && ComponentImageKinds.IsScopeBaseline(row) == baselines,
                rows));
        }

        // ---- files ------------------------------------------------------------------------------

        // EVERY file of the component, whatever region CRT shows - its "Files" page, as AllLinksOf its "Links".
        public static List<ComponentLocalFileEntry> AllFilesOf(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentLocalFiles.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).ToList();
        }

        // The component's whole list of files replaced by `files`, each in its own place.
        public static BoardData WithAllFiles(BoardData board, string? label, IReadOnlyList<ComponentLocalFileEntry> files)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(files);

            if (string.IsNullOrWhiteSpace(label))
            {
                return board;
            }

            bool OfIt(ComponentLocalFileEntry row) => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label);

            List<ComponentLocalFileEntry> before = board.ComponentLocalFiles.Where(OfIt).ToList();
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region: null, before.Select(row => row.BoardLabel));

            List<ComponentLocalFileEntry> rows = files
                .Select(row => new ComponentLocalFileEntry
                {
                    BoardLabel = filedLabel,
                    Name = row.Name?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                })
                .Where(row => row.Name.Length > 0 || row.File.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, row => string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Region?.Trim(), row.Name?.Trim(), row.File?.Trim())))
            {
                return board;
            }

            return board.WithComponentSections(componentLocalFiles: ComponentPaneEdits.ReplaceInScope(board.ComponentLocalFiles, OfIt, rows));
        }

        // ---- shared -----------------------------------------------------------------------------

        // ###########################################################################################
        // The label rows are filed under: the spelling the component's rows of that kind already use
        // (`existing` - typed in other capitals in the table, say), else the component's own spelling,
        // else the one given. Kept as it is (code review, 2026-10-26): every commit rewrote them all
        // under the component's spelling, a change nobody made.
        // ###########################################################################################
        private static string FiledLabel(BoardData board, string? label, string? region, IEnumerable<string?> existing) =>
            existing.Select(own => own?.Trim()).FirstOrDefault(own => !string.IsNullOrEmpty(own)) is { } filed
                ? filed
                : ComponentPaneEdits.RowShown(board, label, region)?.BoardLabel?.Trim() is { Length: > 0 } own
                    ? own
                    : label?.Trim() ?? string.Empty;

        // A label as rows are compared by - its capitals no change, as they are no change to which
        // component a row is (SameLabel).
        private static string LabelKey(string? label) => label?.Trim().ToUpperInvariant() ?? string.Empty;

        private static bool SameLabel(string? rowLabel, string wanted) =>
            string.Equals(rowLabel?.Trim(), wanted, StringComparison.OrdinalIgnoreCase);

        // A row's region is shown in `region` when it is blank (every region) or the same. The ONE rule -
        // the popup (ComponentPopupContent) asks it too (code review, 2026-10-26).
        internal static bool InRegion(string? rowRegion, string? region)
        {
            string own = rowRegion?.Trim() ?? string.Empty;

            return own.Length == 0 || string.Equals(own, region?.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        // A row of the component `label` as shown in `region` - its label, in any capitals, and its region.
        internal static bool InScope(string? rowLabel, string? rowRegion, string? label, string? region) =>
            ComponentPaneEdits.IsOfComponent(rowLabel, label) && ComponentPaneEdits.InRegion(rowRegion, region);

        // A row of the component `label`, whatever its region - its label, in any capitals.
        internal static bool IsOfComponent(string? rowLabel, string? label)
        {
            string wanted = label?.Trim() ?? string.Empty;

            return wanted.Length > 0 && ComponentPaneEdits.SameLabel(rowLabel, wanted);
        }

        internal static string KeyOf(ComponentPinEntry pin) =>
            string.Join('\u001F', ComponentPaneEdits.LabelKey(pin.BoardLabel), pin.Region?.Trim(), pin.Pin?.Trim(), pin.Name?.Trim(), pin.Direction?.Trim());

        internal static string KeyOf(ComponentImageEntry row) =>
            string.Join(
                '\u001F',
                ComponentPaneEdits.LabelKey(row.BoardLabel),
                row.Region?.Trim(),
                row.Pin?.Trim(),
                row.Name?.Trim(),
                row.ExpectedOscilloscopeReading?.Trim(),
                row.File?.Trim(),
                RichNote.NormalisedMarkup(row.Note),
                row.TimeDiv?.Trim(),
                row.VoltsDiv?.Trim(),
                row.TriggerLevelVolts?.Trim(),
                row.MachineState?.Trim());

        // A link and a file by every value they hold - what tells two rows apart when several components are edited
        // together (ComponentBulk), as the keys above do there.
        internal static string KeyOf(ComponentLinkEntry row) =>
            string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Region?.Trim(), row.Name?.Trim(), row.Url?.Trim());

        internal static string KeyOf(ComponentLocalFileEntry row) =>
            string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Region?.Trim(), row.Name?.Trim(), row.File?.Trim());

        internal static string KeyOf(ComponentReplacementEntry row) =>
            string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Region?.Trim(), row.PartNumber?.Trim(), row.Kind?.Trim(), ComponentPaneEdits.NormalisedNote(row.Note));

        private static bool SameRows<T>(IReadOnlyList<T> first, IReadOnlyList<T> second, Func<T, string> key) =>
            first.Count == second.Count && first.Select(key).SequenceEqual(second.Select(key), StringComparer.Ordinal);

        // ###########################################################################################
        // `rows` with the ones `inScope` picks replaced by `replacement`, in the places they held - or at
        // the end, when there was none.
        //
        // *** EACH IN ITS OWN PLACE (code review, 2026-11-02). *** Shipped chips hold their baselines PAL,
        // NTSC, PAL, NTSC. All of them put where the first stood, one value corrected regrouped the lot:
        // "Order of the component pictures changed", to the maintainer and in the History too, regrouped
        // for good at the publish, and a draft that only "reordered" once the value was typed back.
        //
        //   - A row on both sides (by its identity, BoardDraftNaturalKeys) keeps its place, and a place whose
        //     row went is given to the next row that is new - a value of the key corrected is that row still.
        //     With nothing to give it to, the place goes: what follows does not move up past another
        //     region's rows, which would read as a new order too.
        //   - A row new among them goes just before the row it now stands before; past the last, right after
        //     the last place.
        //   - Rows put in ANOTHER ORDER are a reorder, and it is meant: they take the places in turn.
        // ###########################################################################################
        internal static List<T> ReplaceInScope<T>(List<T> rows, Func<T, bool> inScope, IReadOnlyList<T> replacement)
            where T : class
        {
            (bool[] heldKept, bool[] replacementKept, bool sameOrder) = ComponentPaneEdits.RowsKept(rows.Where(inScope).ToList(), replacement);

            var result = new List<T>(rows.Count + replacement.Count);
            int place = 0;
            int next = 0;
            int afterLastPlace = -1;

            foreach (T row in rows)
            {
                if (!inScope(row))
                {
                    result.Add(row);
                    continue;
                }

                if (!sameOrder)
                {
                    if (next < replacement.Count)
                    {
                        result.Add(replacement[next++]);
                    }
                }
                else if (heldKept[place])
                {
                    while (!replacementKept[next])
                    {
                        result.Add(replacement[next++]);
                    }

                    result.Add(replacement[next++]);
                }
                else if (next < replacement.Count && !replacementKept[next])
                {
                    result.Add(replacement[next++]);
                }

                place++;
                afterLastPlace = result.Count;
            }

            result.InsertRange(afterLastPlace < 0 ? result.Count : afterLastPlace, replacement.Skip(next));

            return result;
        }

        // Which rows are on both sides, by identity and once each - two rows of one key pair in order - and
        // whether those stand in the same order on both.
        private static (bool[] HeldKept, bool[] ReplacementKept, bool SameOrder) RowsKept<T>(IReadOnlyList<T> held, IReadOnlyList<T> replacement)
            where T : class
        {
            string[] heldKeys = held.Select(BoardDraftNaturalKeys.ForRow).ToArray();
            string[] replacementKeys = replacement.Select(BoardDraftNaturalKeys.ForRow).ToArray();

            bool[] Kept(string[] keys, string[] others)
            {
                var left = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                foreach (string key in others)
                {
                    left[key] = left.GetValueOrDefault(key) + 1;
                }

                return keys.Select(key =>
                {
                    int count = left.GetValueOrDefault(key);
                    left[key] = count - 1;
                    return count > 0;
                }).ToArray();
            }

            bool[] heldKept = Kept(heldKeys, replacementKeys);
            bool[] replacementKept = Kept(replacementKeys, heldKeys);

            bool sameOrder = heldKeys.Where((_, index) => heldKept[index])
                .SequenceEqual(replacementKeys.Where((_, index) => replacementKept[index]), StringComparer.OrdinalIgnoreCase);

            return (heldKept, replacementKept, sameOrder);
        }
    }
}
