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
        RepairNotes,
    }

    // ###########################################################################################
    // WHAT THE COMPONENT PANE CHANGES IN A DRAFT (CRT 4.0 step 2, 2026-10-25 - the plan: "A
    // component opens in a pane, never in a maximized window ... Basics: a one-line short description
    // with a counter, plus multi-line repair notes ... Pinout ... Replacements").
    //
    // The pane edits a component as it is SHOWN: the region CRT shows, by the component list's own
    // rule (ComponentListBuilder - a row with a blank region is in every region). Every edit is one
    // field, or one section's rows, and arrives here as a transform of the draft's board - applied to
    // the draft held in memory as one step of its history (DraftSession), so it is one Undo away and
    // the table shows it at once.
    //
    //   - A FIELD changes the component's row in the region shown, and nothing else.
    //   - A SECTION (pins, replacement parts, links) replaces the component's rows of that section in
    //     the region shown, where the first of them stood - the other region's rows, and every other
    //     component's, are carried across untouched (ComponentBoardWriter's lesson: "Before assuming a
    //     section can be replaced wholesale, check how it was loaded").
    //   - An edit that changes nothing returns THE SAME board, which DraftSession takes as no step.
    //
    // PURE: takes a board, returns a board, touches no files.
    // ###########################################################################################
    public static class ComponentPaneEdits
    {
        // What a field is called on the pane, and in the draft's history ("Short description of [U8]").
        public static string NameOf(ComponentField field) => field switch
        {
            ComponentField.FriendlyName => "Friendly name",
            ComponentField.TechnicalNameOrValue => "Technical name or value",
            ComponentField.PartNumber => "Part number",
            ComponentField.Category => "Category",
            ComponentField.Package => "Package",
            ComponentField.Description => "Short description",
            ComponentField.RepairNotes => "Repair notes",
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
                ComponentField.RepairNotes => row.RepairNotes,
                _ => throw new ArgumentOutOfRangeException(nameof(field), field, null),
            } ?? string.Empty;
        }

        // ###########################################################################################
        // A field's value as it is kept: trimmed, and on one line - except the repair notes, which
        // keep their lines (each ended by "\n" alone, whatever the clipboard brought).
        // ###########################################################################################
        public static string Normalised(ComponentField field, string? value)
        {
            string text = value ?? string.Empty;

            if (field == ComponentField.RepairNotes)
            {
                return text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
            }

            return text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
        }

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

            if (string.Equals(ComponentPaneEdits.ValueOf(row, field), wanted, StringComparison.Ordinal))
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
                RepairNotes = field == ComponentField.RepairNotes ? wanted : row.RepairNotes,
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

        // Whether the component has a highlight on any picture - what "Show on board" can show.
        public static bool IsMarkedOnAPicture(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            return wanted.Length > 0 && board.ComponentHighlights.Any(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted));
        }

        // ---- pins -------------------------------------------------------------------------------

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
                    Changes = row.Changes?.Trim() ?? string.Empty,
                    Source = row.Source?.Trim() ?? string.Empty,
                })
                .Where(row => row.PartNumber.Length > 0 || row.Kind.Length > 0 || row.Changes.Length > 0 || row.Source.Length > 0)
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

        // The component's links - a link has no region: it is the component's in every region.
        public static List<ComponentLinkEntry> LinksOf(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            return wanted.Length == 0
                ? []
                : board.ComponentLinks.Where(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted)).ToList();
        }

        public static BoardData WithLinks(BoardData board, string? label, IReadOnlyList<ComponentLinkEntry> links)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(links);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return board;
            }

            List<ComponentLinkEntry> before = ComponentPaneEdits.LinksOf(board, label);
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region: null, before.Select(row => row.BoardLabel));

            List<ComponentLinkEntry> rows = links
                .Select(row => new ComponentLinkEntry
                {
                    BoardLabel = filedLabel,
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty,
                })
                .Where(row => row.Name.Length > 0 || row.Url.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, row => string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Name?.Trim(), row.Url?.Trim())))
            {
                return board;
            }

            return board.WithComponentSections(componentLinks: ComponentPaneEdits.ReplaceInScope(
                board.ComponentLinks,
                row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted),
                rows));
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

        // The component's scope baselines in `region`: a pin AND a scope setting.
        public static List<ComponentImageEntry> BaselinesOf(BoardData board, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.ComponentImages
                .Where(row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, region) && ComponentImageKinds.IsScopeBaseline(row))
                .ToList();
        }

        // ###########################################################################################
        // The board with the component's pictures in `region` replaced by `pictures`, where the first
        // of them stood - its scope baselines and every other row left in place (one list holds both).
        // Each filed under the component's label; a row with nothing in it is left out; every value
        // trimmed, a note keeping its lines. The same board when nothing changed.
        // ###########################################################################################
        public static BoardData WithPictures(BoardData board, string? label, string? region, IReadOnlyList<ComponentImageEntry> pictures) =>
            ComponentPaneEdits.WithImages(board, label, region, pictures, baselines: false);

        // ###########################################################################################
        // The same for the component's scope baselines, its pictures left in place. A row handed over
        // that is no baseline by the rule is written as it is - read as a picture from then on: the pane
        // holds back an edit that would do that (a pin cleared, every scope setting cleared), and
        // nothing here drops a row.
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
                    Note = ComponentPaneEdits.Normalised(ComponentField.RepairNotes, row.Note),
                    TimeDiv = row.TimeDiv?.Trim() ?? string.Empty,
                    VoltsDiv = row.VoltsDiv?.Trim() ?? string.Empty,
                    TriggerLevelVolts = row.TriggerLevelVolts?.Trim() ?? string.Empty,
                    MachineState = row.MachineState?.Trim() ?? string.Empty,
                    Source = row.Source?.Trim() ?? string.Empty,
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

        // The component's files - a file has no region: it is the component's in every region.
        public static List<ComponentLocalFileEntry> FilesOf(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            return wanted.Length == 0
                ? []
                : board.ComponentLocalFiles.Where(row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted)).ToList();
        }

        // The same as WithLinks, for the component's files.
        public static BoardData WithFiles(BoardData board, string? label, IReadOnlyList<ComponentLocalFileEntry> files)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(files);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return board;
            }

            List<ComponentLocalFileEntry> before = ComponentPaneEdits.FilesOf(board, label);
            string filedLabel = ComponentPaneEdits.FiledLabel(board, label, region: null, before.Select(row => row.BoardLabel));

            List<ComponentLocalFileEntry> rows = files
                .Select(row => new ComponentLocalFileEntry
                {
                    BoardLabel = filedLabel,
                    Name = row.Name?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty,
                })
                .Where(row => row.Name.Length > 0 || row.File.Length > 0)
                .ToList();

            if (ComponentPaneEdits.SameRows(before, rows, row => string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Name?.Trim(), row.File?.Trim())))
            {
                return board;
            }

            return board.WithComponentSections(componentLocalFiles: ComponentPaneEdits.ReplaceInScope(
                board.ComponentLocalFiles,
                row => ComponentPaneEdits.SameLabel(row.BoardLabel, wanted),
                rows));
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
        // ComponentCopy and the popup (ComponentPopupContent) ask it too (code review, 2026-10-26).
        internal static bool InRegion(string? rowRegion, string? region)
        {
            string own = rowRegion?.Trim() ?? string.Empty;

            return own.Length == 0 || string.Equals(own, region?.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        // A row of the component `label` as shown in `region` - its label, in any capitals, and its region.
        internal static bool InScope(string? rowLabel, string? rowRegion, string? label, string? region)
        {
            string wanted = label?.Trim() ?? string.Empty;

            return wanted.Length > 0 && ComponentPaneEdits.SameLabel(rowLabel, wanted) && ComponentPaneEdits.InRegion(rowRegion, region);
        }

        private static string KeyOf(ComponentPinEntry pin) =>
            string.Join('\u001F', ComponentPaneEdits.LabelKey(pin.BoardLabel), pin.Region?.Trim(), pin.Pin?.Trim(), pin.Name?.Trim(), pin.Direction?.Trim());

        private static string KeyOf(ComponentImageEntry row) =>
            string.Join(
                '\u001F',
                ComponentPaneEdits.LabelKey(row.BoardLabel),
                row.Region?.Trim(),
                row.Pin?.Trim(),
                row.Name?.Trim(),
                row.ExpectedOscilloscopeReading?.Trim(),
                row.File?.Trim(),
                ComponentPaneEdits.Normalised(ComponentField.RepairNotes, row.Note),
                row.TimeDiv?.Trim(),
                row.VoltsDiv?.Trim(),
                row.TriggerLevelVolts?.Trim(),
                row.MachineState?.Trim(),
                row.Source?.Trim());

        private static string KeyOf(ComponentReplacementEntry row) =>
            string.Join('\u001F', ComponentPaneEdits.LabelKey(row.BoardLabel), row.Region?.Trim(), row.PartNumber?.Trim(), row.Kind?.Trim(), row.Changes?.Trim(), row.Source?.Trim());

        private static bool SameRows<T>(IReadOnlyList<T> first, IReadOnlyList<T> second, Func<T, string> key) =>
            first.Count == second.Count && first.Select(key).SequenceEqual(second.Select(key), StringComparer.Ordinal);

        // ###########################################################################################
        // `rows` without the ones `inScope` picks, `replacement` standing where the first of those
        // stood - or at the end, when there was none.
        // ###########################################################################################
        private static List<T> ReplaceInScope<T>(List<T> rows, Func<T, bool> inScope, IReadOnlyList<T> replacement)
        {
            var result = new List<T>(rows.Count + replacement.Count);
            bool placed = false;

            foreach (T row in rows)
            {
                if (!inScope(row))
                {
                    result.Add(row);
                    continue;
                }

                if (!placed)
                {
                    result.AddRange(replacement);
                    placed = true;
                }
            }

            if (!placed)
            {
                result.AddRange(replacement);
            }

            return result;
        }
    }
}
