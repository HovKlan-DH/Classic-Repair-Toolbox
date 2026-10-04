using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE ONE SET OF RULES FOR A BOARD'S ROWS (owner request, 2026-10-02: "integrate the existing
    // validation check into the "Draft" system table view ... either "Error" or "Warning" ... All
    // error should be fixed before submission can be done").
    //
    // There were two checkers until then: the server's SubmissionValidator (errors that refuse a
    // submission, a few warnings) and the app's DataValidator (warnings, written to the log file
    // at launch and nowhere else). Their rules now live HERE, once, and every caller asks this:
    //
    //   - the SERVER (SubmissionValidator), in BoardCheckScope.Submission - exactly the rules,
    //     codes, subjects, messages and order it always had, so what it refuses is unchanged;
    //   - the TABLE (BoardTableDocument, the Drafts tab and the Maintainer tab), in
    //     BoardCheckScope.Everything, marking each problem on its cell;
    //   - the LAUNCH LOG (DataValidator), in Everything too, over every published board.
    //
    // *** AN ERROR IN THE TABLE IS A REFUSAL AT SUBMIT, AND NOTHING ELSE IS. *** Everything adds
    // to Submission only (a) the server's file-name rules, which it applies in another class
    // (SubmissionPathRules.IsSafelyShaped, SubmissionFileRules.TryCheckName - called here, not
    // restated), and (b) WARNINGS. BoardDataChecksTests holds that: every error Everything finds
    // for a defect is one the server's own checks find for it too.
    //
    // *** WHERE A PROBLEM IS, IS AN ENTRY INDEX. *** Index N on a sheet is entry N of that sheet's
    // section - the order BoardWorkbookSchema maps and BuildRows writes - so the table maps it back
    // to the row it built that entry from.
    //
    // Pure: rows and a file lookup in, problems out. The lookup is optional - with none, the
    // existence and spelling of files is simply not checked (the Maintainer tab's table, whose
    // files the server checked when the submission arrived).
    // ###########################################################################################
    public static class BoardDataChecks
    {
        // ###########################################################################################
        // The settings CRT's oscilloscope support knows - what DataValidator logged anything else
        // as a problem against since before submissions existed. Empty is allowed: the columns are
        // optional.
        // ###########################################################################################
        private static readonly HashSet<string> AllowedTimeDivValues = new(StringComparer.OrdinalIgnoreCase)
        {
            "2nS", "5nS", "10nS", "20nS", "50nS", "100nS", "200nS", "500nS",
            "1uS", "2uS", "5uS", "10uS", "20uS", "50uS", "100uS", "200uS", "500uS",
            "1mS", "2mS", "5mS", "10mS", "20mS", "50mS", "100mS", "200mS", "500mS",
            "1S", "2S", "5S", "10S", "20S", "50S", "100S", "200S", "500S", "1000S"
        };

        private static readonly HashSet<string> AllowedVoltsDivValues = new(StringComparer.OrdinalIgnoreCase)
        {
            "5mV", "10mV", "20mV", "50mV", "100mV", "200mV", "500mV",
            "1V", "2V", "5V", "10V", "20V", "50V", "100V"
        };

        private static readonly Regex TriggerLevelRegex = new(
            @"^-?\d+(?:\.\d+)?[Vv]$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static IReadOnlyList<BoardDataProblem> Check(BoardCheckRows rows, IBoardFileLookup? files, BoardCheckScope scope)
        {
            ArgumentNullException.ThrowIfNull(rows);

            var context = new Context(rows, files, scope);

            // The server's rules, in the server's order (SubmissionValidator.Validate).
            BoardDataChecks.CheckSchematics(context);
            BoardDataChecks.CheckComponents(context);
            BoardDataChecks.CheckComponentImages(context);
            BoardDataChecks.CheckHighlights(context);
            BoardDataChecks.CheckLocalFiles(context);
            BoardDataChecks.CheckLinks(context);

            // The warnings DataValidator used to log on its own.
            if (scope == BoardCheckScope.Everything)
            {
                BoardDataChecks.CheckOscilloscopeSettings(context);
                BoardDataChecks.CheckOrphanFilesAndLinks(context);
                BoardDataChecks.CheckHighlightCoverage(context);
                BoardDataChecks.CheckPartNumbers(context);
                BoardDataChecks.CheckDuplicateRows(context);
            }

            return context.Problems;
        }

        // ###########################################################################################
        // "Board label, Region, Pin and Name" - columns as a sentence names them.
        // ###########################################################################################
        internal static string ListOf(IReadOnlyList<string> names) => names.Count switch
        {
            0 => string.Empty,
            1 => names[0],
            _ => $"{string.Join(", ", names.Take(names.Count - 1))} and {names[^1]}"
        };

        // ###########################################################################################
        // The highest level among problems - None when there are none.
        // ###########################################################################################
        public static BoardProblemLevel Worst(IEnumerable<BoardDataProblem> problems) =>
            problems.Select(problem => problem.Level).DefaultIfEmpty(BoardProblemLevel.None).Max();

        // -------------------------------------------------------------------------------------
        // The server's rules - see SubmissionValidator for the reasoning behind each.
        // -------------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE FIRST ROW OF A DUPLICATE IS MARKED TOO, IN THE TABLE (owner decision, 2026-10-03:
        // "it should show all rows, and not only last, as the maintainer/contributor then has some
        // context"). *** The server refuses on the LATER rows, as it always has, so Submission's
        // findings are unchanged; Everything adds the same error on the first row of each set, so a
        // duplicate is never one marked row beside a twin that looks fine. Same code, so the table
        // still finds no error the server does not. Applies to schematics and components - the two
        // sheets where a duplicate is the server's error (the other sheets' duplicates are warnings
        // on every row already, CheckDuplicateRows).
        // ###########################################################################################
        private static void CheckSchematics(Context context)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var firstRowOf = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var duplicated = new SortedSet<int>();
            const string sheet = BoardWorkbookSchema.SheetBoardSchematics;

            for (int i = 0; i < context.Rows.Schematics.Count; i++)
            {
                BoardSchematicEntry schematic = context.Rows.Schematics[i];

                if (string.IsNullOrWhiteSpace(schematic.SchematicName))
                {
                    context.Error("schematic.unnamed", string.Empty, "A schematic row has no name.",
                        sheet, i, BoardWorkbookSchema.ColSchematicName);

                    continue;
                }

                // Case-insensitive: two schematics differing only by capitalisation would be
                // indistinguishable to a human, and highlights reference them by name.
                if (!seen.Add(schematic.SchematicName))
                {
                    context.Error(
                        "schematic.duplicate",
                        schematic.SchematicName,
                        $"More than one schematic is named [{schematic.SchematicName}].",
                        sheet, i, BoardWorkbookSchema.ColSchematicName);

                    duplicated.Add(firstRowOf[schematic.SchematicName]);
                }
                else
                {
                    firstRowOf[schematic.SchematicName] = i;
                }

                context.CheckFileReference(
                    schematic.SchematicImageFile,
                    $"schematic [{schematic.SchematicName}]",
                    sheet, i, BoardWorkbookSchema.ColSchematicImageFile);
            }

            if (context.Scope != BoardCheckScope.Everything)
                return;

            foreach (int first in duplicated)
            {
                string name = context.Rows.Schematics[first].SchematicName;

                context.Error(
                    "schematic.duplicate",
                    name,
                    $"More than one schematic is named [{name}] - this is the first of them, the " +
                    "other(s) further down.",
                    sheet, first, BoardWorkbookSchema.ColSchematicName);
            }
        }

        private static void CheckComponents(Context context)
        {
            var seen = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var firstRowOf = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var duplicated = new SortedSet<int>();
            const string sheet = BoardWorkbookSchema.SheetComponents;

            for (int i = 0; i < context.Rows.Components.Count; i++)
            {
                ComponentEntry component = context.Rows.Components[i];

                if (string.IsNullOrWhiteSpace(component.BoardLabel))
                {
                    context.Error(
                        "component.unlabelled",
                        string.Empty,
                        "A component row has no board label. The label is how everything else " +
                        "refers to it, so a row without one cannot be used.",
                        sheet, i, BoardWorkbookSchema.ColBoardLabel);

                    continue;
                }

                // LABEL + REGION, not the label alone - a regional variant is one row per region
                // (SubmissionValidator's header has the history). A blank region is its own bucket.
                string region = component.Region?.Trim() ?? string.Empty;

                string key = region.Length == 0
                    ? component.BoardLabel.Trim()
                    : component.BoardLabel.Trim() + BoardDraftNaturalKeys.Separator + region;

                if (seen.TryGetValue(key, out string? first))
                {
                    string where = region.Length == 0
                        ? "with no region"
                        : $"in region [{region}]";

                    context.Error(
                        "component.duplicate_label",
                        component.BoardLabel,
                        $"More than one component is labelled [{component.BoardLabel}] {where} " +
                        $"(also [{first}]). Board labels must be unique within a region.",
                        sheet, i, BoardWorkbookSchema.ColBoardLabel);

                    duplicated.Add(firstRowOf[key]);
                }
                else
                {
                    seen[key] = component.BoardLabel.Trim();
                    firstRowOf[key] = i;
                }
            }

            if (context.Scope != BoardCheckScope.Everything)
                return;

            foreach (int first in duplicated)
            {
                ComponentEntry component = context.Rows.Components[first];
                string region = component.Region?.Trim() ?? string.Empty;
                string where = region.Length == 0 ? "with no region" : $"in region [{region}]";

                context.Error(
                    "component.duplicate_label",
                    component.BoardLabel,
                    $"More than one component is labelled [{component.BoardLabel.Trim()}] {where} - this " +
                    "is the first of them, the other(s) further down. Board labels must be unique within a region.",
                    sheet, first, BoardWorkbookSchema.ColBoardLabel);
            }
        }

        // ###########################################################################################
        // A component image row with a note and no file IS a note, and valid: the "Pinout" rows that
        // give only a compatible part number (69 of them across the shipped boards). A row with
        // neither declares nothing. The ONE answer to "does this row need a file" - the component
        // editor asks it too (ContributionPackaging.ValidateComponentImageFile), so the editor
        // cannot refuse a row the table and the server accept (owner report, 2026-10-02).
        // ###########################################################################################
        public static bool IsNoteOnlyImage(string? file, string? note) =>
            string.IsNullOrWhiteSpace(file) && !string.IsNullOrWhiteSpace(note);

        private static void CheckComponentImages(Context context)
        {
            var componentLabels = new HashSet<string>(
                context.Rows.Components.Select(component => component.BoardLabel),
                StringComparer.OrdinalIgnoreCase);

            const string sheet = BoardWorkbookSchema.SheetComponentImages;

            for (int i = 0; i < context.Rows.ComponentImages.Count; i++)
            {
                ComponentImageEntry image = context.Rows.ComponentImages[i];

                // A row with a Note and no File is a note, and valid; a row with neither declares
                // nothing and is still reported.
                if (!BoardDataChecks.IsNoteOnlyImage(image.File, image.Note))
                {
                    context.CheckFileReference(
                        image.File,
                        $"component image for [{image.BoardLabel}]",
                        sheet, i, BoardWorkbookSchema.ColFile);
                }

                if (!string.IsNullOrWhiteSpace(image.BoardLabel) &&
                    !componentLabels.Contains(image.BoardLabel))
                {
                    context.Warning(
                        "image.orphan",
                        image.BoardLabel,
                        $"An image references component [{image.BoardLabel}], which is not in this " +
                        "submission. It will not be shown anywhere.",
                        sheet, i, BoardWorkbookSchema.ColBoardLabel);
                }
            }
        }

        // Highlights live in the JSON beside the workbook, in no sheet - so they have no place in
        // the table and are told about apart from it.
        private static void CheckHighlights(Context context)
        {
            var schematicNames = new HashSet<string>(
                context.Rows.Schematics.Select(schematic => schematic.SchematicName),
                StringComparer.OrdinalIgnoreCase);

            foreach (ComponentHighlightEntry highlight in context.Rows.ComponentHighlights)
            {
                string subject = $"{highlight.SchematicName} / {highlight.BoardLabel}";

                if (!schematicNames.Contains(highlight.SchematicName))
                {
                    context.Error(
                        "highlight.unknown_schematic",
                        subject,
                        $"A highlight names schematic [{highlight.SchematicName}], which is not in " +
                        "this submission.",
                        sheet: null, index: -1, column: null);
                }

                if (!BoardDataChecks.TryParseCoordinate(highlight.X, out double x) ||
                    !BoardDataChecks.TryParseCoordinate(highlight.Y, out double y) ||
                    !BoardDataChecks.TryParseCoordinate(highlight.Width, out double width) ||
                    !BoardDataChecks.TryParseCoordinate(highlight.Height, out double height))
                {
                    context.Error(
                        "highlight.unparseable",
                        subject,
                        $"A highlight has coordinates that cannot be read " +
                        $"(x=[{highlight.X}] y=[{highlight.Y}] w=[{highlight.Width}] h=[{highlight.Height}]). " +
                        "Numbers must use a dot as the decimal separator.",
                        sheet: null, index: -1, column: null);

                    continue;
                }

                if (width <= 0 || height <= 0)
                {
                    context.Error(
                        "highlight.degenerate",
                        subject,
                        $"A highlight has no area (width {width}, height {height}). It would be " +
                        "invisible and could not be adjusted.",
                        sheet: null, index: -1, column: null);
                }

                if (x < 0 || y < 0)
                {
                    context.Error(
                        "highlight.negative_origin",
                        subject,
                        $"A highlight starts outside the image (x {x}, y {y}).",
                        sheet: null, index: -1, column: null);
                }
            }
        }

        private static void CheckLocalFiles(Context context)
        {
            for (int i = 0; i < context.Rows.ComponentLocalFiles.Count; i++)
            {
                ComponentLocalFileEntry file = context.Rows.ComponentLocalFiles[i];

                context.CheckFileReference(
                    file.File, $"file for [{file.BoardLabel}]",
                    BoardWorkbookSchema.SheetComponentLocalFiles, i, BoardWorkbookSchema.ColFile);
            }

            for (int i = 0; i < context.Rows.BoardLocalFiles.Count; i++)
            {
                context.CheckFileReference(
                    context.Rows.BoardLocalFiles[i].File, "board file",
                    BoardWorkbookSchema.SheetBoardLocalFiles, i, BoardWorkbookSchema.ColFile);
            }
        }

        private static void CheckLinks(Context context)
        {
            for (int i = 0; i < context.Rows.ComponentLinks.Count; i++)
            {
                ComponentLinkEntry link = context.Rows.ComponentLinks[i];

                BoardDataChecks.CheckLink(
                    context, link.Url, $"link for [{link.BoardLabel}]",
                    BoardWorkbookSchema.SheetComponentLinks, i);
            }

            // A board link is scoped by Category rather than by a component label.
            for (int i = 0; i < context.Rows.BoardLinks.Count; i++)
            {
                BoardDataChecks.CheckLink(
                    context, context.Rows.BoardLinks[i].Url, "board link",
                    BoardWorkbookSchema.SheetBoardLinks, i);
            }
        }

        // A link must be http or https - see SubmissionValidator: these are opened through
        // ExternalTargetLauncher, which admits nothing else.
        private static void CheckLink(Context context, string? link, string description, string sheet, int index)
        {
            if (string.IsNullOrWhiteSpace(link))
                return;

            bool isWebLink = Uri.TryCreate(link, UriKind.Absolute, out Uri? uri)
                && (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));

            if (!isWebLink)
            {
                context.Error(
                    "link.not_web",
                    link,
                    $"A {description} is [{link}], which is not an http:// or https:// address. " +
                    "CRT will not open anything else.",
                    sheet, index, BoardWorkbookSchema.ColUrl);
            }
        }

        // -------------------------------------------------------------------------------------
        // The local warnings - what DataValidator logged at launch, now shown where the row is.
        // -------------------------------------------------------------------------------------

        private static void CheckOscilloscopeSettings(Context context)
        {
            const string sheet = BoardWorkbookSchema.SheetComponentImages;

            for (int i = 0; i < context.Rows.ComponentImages.Count; i++)
            {
                ComponentImageEntry image = context.Rows.ComponentImages[i];

                string timeDiv = (image.TimeDiv ?? string.Empty).Trim();
                if (timeDiv.Length > 0 && !BoardDataChecks.AllowedTimeDivValues.Contains(timeDiv))
                {
                    context.Warning(
                        "image.time_div",
                        timeDiv,
                        $"T/DIV [{timeDiv}] is not a setting CRT knows. It is written like 500nS, 20uS, 1mS or 2S.",
                        sheet, i, BoardWorkbookSchema.ColTimeDiv);
                }

                string voltsDiv = (image.VoltsDiv ?? string.Empty).Trim();
                if (voltsDiv.Length > 0 && !BoardDataChecks.AllowedVoltsDivValues.Contains(voltsDiv))
                {
                    context.Warning(
                        "image.volts_div",
                        voltsDiv,
                        $"V/DIV [{voltsDiv}] is not a setting CRT knows. It is written like 50mV, 500mV, 1V or 5V.",
                        sheet, i, BoardWorkbookSchema.ColVoltsDiv);
                }

                string triggerLevel = (image.TriggerLevelVolts ?? string.Empty).Trim();
                if (triggerLevel.Length > 0 && !BoardDataChecks.TriggerLevelRegex.IsMatch(triggerLevel))
                {
                    context.Warning(
                        "image.trigger_level",
                        triggerLevel,
                        $"T.LVL [{triggerLevel}] is not a voltage CRT can read. It is written like 1.4V or -2V.",
                        sheet, i, BoardWorkbookSchema.ColTriggerLevelVolts);
                }
            }
        }

        // The image sheet's orphans are the server's own warning (image.orphan, above); files and
        // links naming a component that does not exist were only ever in the log.
        private static void CheckOrphanFilesAndLinks(Context context)
        {
            HashSet<string> labels = context.ComponentLabels;

            for (int i = 0; i < context.Rows.ComponentLocalFiles.Count; i++)
            {
                string label = (context.Rows.ComponentLocalFiles[i].BoardLabel ?? string.Empty).Trim();

                if (label.Length > 0 && !labels.Contains(label))
                {
                    context.Warning(
                        "file.orphan",
                        label,
                        $"This file is for component [{label}], which is not in the Components sheet. It will not be shown anywhere.",
                        BoardWorkbookSchema.SheetComponentLocalFiles, i, BoardWorkbookSchema.ColBoardLabel);
                }
            }

            for (int i = 0; i < context.Rows.ComponentLinks.Count; i++)
            {
                string label = (context.Rows.ComponentLinks[i].BoardLabel ?? string.Empty).Trim();

                if (label.Length > 0 && !labels.Contains(label))
                {
                    context.Warning(
                        "link.orphan",
                        label,
                        $"This link is for component [{label}], which is not in the Components sheet. It will not be shown anywhere.",
                        BoardWorkbookSchema.SheetComponentLinks, i, BoardWorkbookSchema.ColBoardLabel);
                }
            }
        }

        // ###########################################################################################
        // Components and highlights that do not meet: a component nothing marks on any schematic,
        // and a highlight for a component that is not there. Highlights are drawn in the label
        // editor, not typed into a sheet, so a component's own row is where its missing highlight
        // is told; a stray highlight has no row at all.
        // ###########################################################################################
        private static void CheckHighlightCoverage(Context context)
        {
            var highlighted = new HashSet<string>(
                context.Rows.ComponentHighlights
                    .Select(highlight => (highlight.BoardLabel ?? string.Empty).Trim())
                    .Where(label => label.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < context.Rows.Components.Count; i++)
            {
                string label = (context.Rows.Components[i].BoardLabel ?? string.Empty).Trim();

                if (label.Length > 0 && !highlighted.Contains(label))
                {
                    context.Warning(
                        "component.no_highlight",
                        label,
                        $"Component [{label}] is not marked on any schematic, so CRT cannot point it out on the board. " +
                        "Mark it with the label editor on the Schematics tab.",
                        BoardWorkbookSchema.SheetComponents, i, BoardWorkbookSchema.ColBoardLabel);
                }
            }

            HashSet<string> labels = context.ComponentLabels;

            foreach (ComponentHighlightEntry highlight in context.Rows.ComponentHighlights)
            {
                string label = (highlight.BoardLabel ?? string.Empty).Trim();

                if (label.Length > 0 && !labels.Contains(label))
                {
                    context.Warning(
                        "highlight.orphan",
                        $"{highlight.SchematicName} / {label}",
                        $"Schematic [{highlight.SchematicName}] marks component [{label}], which is not in the Components sheet.",
                        sheet: null, index: -1, column: null);
                }
            }
        }

        // ###########################################################################################
        // One part number, two different chips: a part number names ONE chip, so the same one on
        // components with different technical names is almost certainly a copy-and-paste slip (the
        // C64 250407's PLA once carried the 4066's). Every row involved is marked, naming the others.
        // ###########################################################################################
        private static void CheckPartNumbers(Context context)
        {
            var byPartNumber = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < context.Rows.Components.Count; i++)
            {
                ComponentEntry component = context.Rows.Components[i];
                string partNumber = (component.PartNumber ?? string.Empty).Trim();
                string technicalName = (component.TechnicalNameOrValue ?? string.Empty).Trim();

                if (partNumber.Length == 0 || technicalName.Length == 0)
                    continue;

                if (!byPartNumber.TryGetValue(partNumber, out List<int>? indexes))
                {
                    indexes = [];
                    byPartNumber[partNumber] = indexes;
                }

                indexes.Add(i);
            }

            foreach ((string partNumber, List<int> indexes) in byPartNumber)
            {
                List<string> names = indexes
                    .Select(i => context.Rows.Components[i].TechnicalNameOrValue.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (names.Count < 2)
                    continue;

                foreach (int i in indexes)
                {
                    ComponentEntry component = context.Rows.Components[i];
                    string own = component.TechnicalNameOrValue.Trim();

                    IEnumerable<string> others = indexes
                        .Where(other => other != i)
                        .Select(other => context.Rows.Components[other])
                        .Where(other => !string.Equals(other.TechnicalNameOrValue.Trim(), own, StringComparison.OrdinalIgnoreCase))
                        .Select(other => $"{other.BoardLabel.Trim()} ({other.TechnicalNameOrValue.Trim()})");

                    context.Warning(
                        "component.part_number_mixed",
                        partNumber,
                        $"Part-number [{partNumber}] is also used for {string.Join(", ", others)}. " +
                        "One part-number names one kind of chip - check it was not copied from another row.",
                        BoardWorkbookSchema.SheetComponents, i, BoardWorkbookSchema.ColPartNumber);
                }
            }
        }

        // ###########################################################################################
        // *** ROWS WITH THE SAME IDENTITY (owner request, 2026-10-03: "Should flagged now be treated
        // as warnings?"; cases agreed with the project owner). *** Two rows on a sheet with the
        // same natural key are both saved and published, but BoardDataDiffer pairs only the FIRST
        // with the published data - so a change in the others is counted nowhere and a maintainer
        // is shown nothing of it. The table drew such rows violet ("Flagged") until then; it is a
        // warning now, on EVERY row of the set, on the first identity column. The identity is
        // BoardDraftNaturalKeys' - the differ's own.
        //
        // Components and Board schematics are not here: a duplicate there is the server's ERROR
        // (component.duplicate_label, schematic.duplicate, above), and a row is not told twice.
        // ###########################################################################################
        private static void CheckDuplicateRows(Context context)
        {
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetComponentImages, context.Rows.ComponentImages);
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetComponentLocalFiles, context.Rows.ComponentLocalFiles);
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetComponentLinks, context.Rows.ComponentLinks);
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetBoardLocalFiles, context.Rows.BoardLocalFiles);
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetBoardLinks, context.Rows.BoardLinks);
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetCredits, context.Rows.Credits ?? []);
            BoardDataChecks.WarnDuplicates(context, BoardWorkbookSchema.SheetKiCadImportantSignals, context.Rows.ImportantSignals ?? []);
        }

        private static void WarnDuplicates<T>(Context context, string sheet, IReadOnlyList<T> entries)
            where T : notnull
        {
            var keys = entries.Select(entry => BoardDraftNaturalKeys.ForRow(entry)).ToList();

            Dictionary<string, int> countByKey = keys
                .GroupBy(key => key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

            IReadOnlyList<string> columns = BoardDraftNaturalKeys.ColumnsOf(sheet);
            string identity = BoardDataChecks.ListOf(columns);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < entries.Count; i++)
            {
                int count = countByKey[keys[i]];
                if (count < 2)
                    continue;

                string message = seen.Add(keys[i])
                    ? $"Another row on this sheet has the same {identity}. Only this one, the first, is compared with the " +
                      $"published data, so a change in the {(count == 2 ? "other" : "others")} is not shown to a maintainer. " +
                      "Make them differ, or delete one."
                    : $"Another row on this sheet has the same {identity}. Only the first is compared with the published " +
                      "data, so a change in this one is not shown to a maintainer. Make them differ, or delete one.";

                context.Warning("row.duplicate", BoardDraftNaturalKeys.Describe(keys[i]), message, sheet, i, columns[0]);
            }
        }

        // ###########################################################################################
        // INVARIANT CULTURE ONLY - see SubmissionValidator: "1,5" is a malformed value, not 1.5.
        // ###########################################################################################
        private static bool TryParseCoordinate(string? value, out double result)
        {
            result = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        // ###########################################################################################
        // One run's state: the rows, the lookup, and the problems found so far.
        // ###########################################################################################
        private sealed class Context(BoardCheckRows rows, IBoardFileLookup? files, BoardCheckScope scope)
        {
            private HashSet<string>? thisComponentLabels;

            public BoardCheckRows Rows { get; } = rows;

            public BoardCheckScope Scope { get; } = scope;

            public List<BoardDataProblem> Problems { get; } = [];

            // The Components sheet's labels, trimmed and ignoring case - the local rules' view.
            public HashSet<string> ComponentLabels => this.thisComponentLabels ??= new HashSet<string>(
                this.Rows.Components
                    .Select(component => (component.BoardLabel ?? string.Empty).Trim())
                    .Where(label => label.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            public void Error(string code, string subject, string message, string? sheet, int index, string? column) =>
                this.Problems.Add(new BoardDataProblem(BoardProblemLevel.Error, code, subject, message, sheet, index, column));

            public void Warning(string code, string subject, string message, string? sheet, int index, string? column) =>
                this.Problems.Add(new BoardDataProblem(BoardProblemLevel.Warning, code, subject, message, sheet, index, column));

            // ###########################################################################################
            // A file a row names: named at all, then - in Everything - shaped and typed as the
            // server will accept, then there and spelled right. The server's own checks of shape and
            // type are separate classes; they are CALLED here, never restated. One problem per cell:
            // a path that is wrongly shaped is not also reported missing.
            // ###########################################################################################
            public void CheckFileReference(string? reference, string description, string sheet, int index, string column)
            {
                if (string.IsNullOrWhiteSpace(reference))
                {
                    this.Error("file.unreferenced", description, $"The {description} does not name a file.", sheet, index, column);
                    return;
                }

                if (this.Scope == BoardCheckScope.Everything)
                {
                    if (!SubmissionPathRules.IsSafelyShaped(reference, out string shapeReason))
                    {
                        this.Error("path.rejected", reference, shapeReason, sheet, index, column);
                        return;
                    }

                    if (!SubmissionFileRules.TryCheckName(reference, out string nameCode, out string nameReason))
                    {
                        this.Error(nameCode, reference, nameReason, sheet, index, column);
                        return;
                    }
                }

                if (files is null)
                    return;

                BoardFileLookupResult found = files.Check(reference);

                if (found.State == BoardFileState.CaseDiffers)
                {
                    this.Error(
                        "file.case_mismatch",
                        reference,
                        $"The {description} names [{reference}], but the file {files.Where} is spelled " +
                        $"[{found.Actual}]. Capitalisation matters on the server even though it does not on " +
                        "Windows, so this would work for you and fail for everyone else.",
                        sheet, index, column);

                    return;
                }

                if (found.State == BoardFileState.Missing)
                {
                    this.Error(
                        "file.missing",
                        reference,
                        $"The {description} names [{reference}], which is not {files.Where}.",
                        sheet, index, column);
                }
            }
        }
    }

    // ###########################################################################################
    // The sections the checks read - BoardData's, or a submission's (SubmissionRows), which carry
    // the same lists under the same names.
    //
    // Credits and Important signals are read by one local rule only - duplicate rows (2026-10-03)
    // - so they are optional, and a submission's rows (which the server checks in Submission,
    // where that rule does not run) leave them out.
    // ###########################################################################################
    public sealed record BoardCheckRows(
        IReadOnlyList<BoardSchematicEntry> Schematics,
        IReadOnlyList<ComponentEntry> Components,
        IReadOnlyList<ComponentImageEntry> ComponentImages,
        IReadOnlyList<ComponentHighlightEntry> ComponentHighlights,
        IReadOnlyList<ComponentLocalFileEntry> ComponentLocalFiles,
        IReadOnlyList<BoardLocalFileEntry> BoardLocalFiles,
        IReadOnlyList<ComponentLinkEntry> ComponentLinks,
        IReadOnlyList<BoardLinkEntry> BoardLinks,
        IReadOnlyList<CreditEntry>? Credits = null,
        IReadOnlyList<KiCadImportantSignalEntry>? ImportantSignals = null)
    {
        public static BoardCheckRows From(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return new BoardCheckRows(
                board.Schematics, board.Components, board.ComponentImages, board.ComponentHighlights,
                board.ComponentLocalFiles, board.BoardLocalFiles, board.ComponentLinks, board.BoardLinks,
                board.Credits, board.KiCadImportantSignals);
        }

        public static BoardCheckRows From(SubmissionRows rows)
        {
            ArgumentNullException.ThrowIfNull(rows);

            return new BoardCheckRows(
                rows.Schematics, rows.Components, rows.ComponentImages, rows.ComponentHighlights,
                rows.ComponentLocalFiles, rows.BoardLocalFiles, rows.ComponentLinks, rows.BoardLinks);
        }
    }
}
