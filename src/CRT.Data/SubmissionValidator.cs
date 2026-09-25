using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Checks a submission's ROWS before any human sees it. Pure: rows and a set of known file
    // paths in, findings out. It opens nothing and reads no images, so every rule is a unit test.
    //
    // *** THIS IS THE HIGHEST-LEVERAGE WORK IN THE WHOLE PLAN. *** NewContributeStrategy.md says
    // so outright, and the reason is arithmetic: the maintainer is one volunteer. Every submission
    // rejected automatically with a clear explanation is reviewer time not spent, and every bad
    // submission that reaches a human costs far more than the contributor saved by not checking.
    // A rejection here is also FASTER for the contributor than a review queue.
    //
    // HARD FAILURES ARE REJECTED AUTOMATICALLY AND NEVER QUEUED. Warnings are attached for the
    // reviewer to weigh. The line between them is "could a reasonable contributor have meant
    // this?" - a highlight outside its image is a mistake in any reading, while an unusually large
    // number of components is merely worth a glance.
    //
    // WHAT IS NOT HERE: anything needing the file bytes. Image decoding is done by the server,
    // which has the blobs; this class takes the SET OF PATHS the submission provides and checks
    // that every reference resolves within it. That split keeps the rules testable without
    // fixtures, and it is why the file-reference checks below are about names rather than content.
    //
    // *** FILE NAMES ARE COMPARED CASE-SENSITIVELY. *** This is stated as a trap in the strategy
    // document and is easy to get wrong in the name of being helpful. From Phase 3 the data tree
    // is read on Linux, where CRT.Data's exact-case File.Exists means a row naming "foo.pdf" when
    // the file is "foo.PDF" loads fine on the contributor's Windows machine and fails only on the
    // server, after publication. Making this comparison case-insensitive would accept data that is
    // genuinely broken and leave Windows clients disagreeing with the server about which file is
    // meant. Do not "fix" it.
    // ###########################################################################################
    public static class SubmissionValidator
    {
        // Above this, a submission is worth a human glance before it is merged - not wrong, but
        // far outside what any shipped board carries (the largest is a few hundred components).
        public const int ImplausibleComponentCount = 2000;

        public const int ImplausibleSchematicCount = 100;

        // ###########################################################################################
        // Validates the rows of a submission.
        //
        // suppliedPaths is every file path the submission provides, compared CASE-SENSITIVELY.
        // ###########################################################################################
        public static IReadOnlyList<ValidationFinding> Validate(
            SubmissionManifest manifest,
            IReadOnlyCollection<string> suppliedPaths)
        {
            ArgumentNullException.ThrowIfNull(manifest);
            ArgumentNullException.ThrowIfNull(suppliedPaths);

            var findings = new List<ValidationFinding>();

            // Ordinal: see the header. This is the comparison the server's filesystem will make.
            var files = new HashSet<string>(suppliedPaths, StringComparer.Ordinal);

            // ...and a case-insensitive copy, used ONLY to give a better message when a reference
            // fails. Telling someone "the file is there but spelled differently" is far more
            // useful than "file not found", and it is the mistake they will actually have made.
            var filesIgnoringCase = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (string path in suppliedPaths)
                filesIgnoringCase.TryAdd(path, path);

            SubmissionValidator.ValidateIdentity(manifest, findings);
            SubmissionValidator.ValidateSchematics(manifest, files, filesIgnoringCase, findings);
            SubmissionValidator.ValidateComponents(manifest, findings);
            SubmissionValidator.ValidateComponentImages(manifest, files, filesIgnoringCase, findings);
            SubmissionValidator.ValidateHighlights(manifest, findings);
            SubmissionValidator.ValidateLocalFiles(manifest, files, filesIgnoringCase, findings);
            SubmissionValidator.ValidateLinks(manifest, findings);
            SubmissionValidator.ValidateScale(manifest, findings);

            return findings;
        }

        // ###########################################################################################
        // True when nothing found is fatal. A submission with only warnings is queued; one with any
        // error never is.
        // ###########################################################################################
        public static bool CanBeQueued(IReadOnlyList<ValidationFinding> findings)
        {
            ArgumentNullException.ThrowIfNull(findings);

            return !findings.Any(finding => finding.Severity == ValidationSeverity.Error);
        }

        // -------------------------------------------------------------------------------------
        // Identity.
        // -------------------------------------------------------------------------------------

        private static void ValidateIdentity(SubmissionManifest manifest, List<ValidationFinding> findings)
        {
            if (manifest.FormatVersion != SubmissionFormat.CurrentVersion)
            {
                findings.Add(SubmissionValidator.Error(
                    "format.unsupported",
                    string.Empty,
                    $"This submission uses format version {manifest.FormatVersion}, but this server " +
                    $"understands version {SubmissionFormat.CurrentVersion}. Update CRT and submit again."));
            }

            // ###########################################################################################
            // THE SystemId IS REQUIRED AND MUST BE THE RIGHT SHAPE.
            //
            // It is "Manufacturer/Hardware/Board" and the client can always compute it from what
            // the contributor typed, so an absent or malformed one is a broken client rather than
            // a new system - a NEW system carries an id too, and is new because no `systems` row
            // holds that id yet.
            //
            // This matters more than an ordinary field check: the value is a DATABASE PRIMARY KEY
            // (and a foreign key from `submissions`), and it is used to resolve a folder in the
            // published tree. A malformed one is either a broken client or an attempt to attach a
            // submission to something it does not belong to.
            //
            // Shape is all that can be decided here. Whether the id names a system that EXISTS,
            // and whether this contributor may submit to it, are database questions the server
            // answers separately.
            // ###########################################################################################
            if (!SystemDescriptorRules.IsValidSystemId(manifest.SystemId))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.system_id_malformed",
                    string.Empty,
                    "The submission does not carry a valid system identifier. " +
                    "This is a fault in the submitting application rather than in your data."));
            }

            // The id is built from these three, so a disagreement means the client assembled the
            // manifest inconsistently - and the id is what everything keys off, so the parts would
            // silently be the wrong ones on any screen that showed them.
            if (SystemDescriptorRules.IsValidSystemId(manifest.SystemId)
                && !string.Equals(
                    manifest.SystemId,
                    SystemDescriptorRules.BuildSystemId(manifest.Manufacturer, manifest.Hardware, manifest.Board),
                    StringComparison.Ordinal))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.system_id_mismatch",
                    string.Empty,
                    "The submission's system identifier does not match the manufacturer, hardware " +
                    "and board it names. This is a fault in the submitting application."));
            }

            // ###########################################################################################
            // *** THE PARTS MUST ALREADY BE IN THEIR CANONICAL FORM (security review, 2026-09-25). ***
            //
            // BuildSystemId trims and collapses whitespace before joining, so "Commodore" followed
            // by a thousand spaces produced a VALID, MATCHING id - and the raw thousand-character
            // part then went into systems.manufacturer, a VARCHAR(100), failing the insert with a
            // 500. The client builds these from folder names, which are already canonical, so
            // requiring equality refuses nothing real.
            // ###########################################################################################
            if (!SubmissionValidator.IsCanonical(manifest.Manufacturer) ||
                !SubmissionValidator.IsCanonical(manifest.Hardware) ||
                !SubmissionValidator.IsCanonical(manifest.Board))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.parts_not_canonical",
                    string.Empty,
                    "The manufacturer, hardware or board name carries extra spaces. " +
                    "This is a fault in the submitting application rather than in your data."));
            }

            // ###########################################################################################
            // *** A BOARD MAY NOT SIT WHERE THE SHARED FOLDERS DO. *** A manufacturer called
            // "Generic shared files", or hardware called "Shared files", would make this system's
            // own folder the same place as a shared one - and a submission may change its own folder
            // freely, so that would turn the shared-folder rules in SubmissionFileScope inside out.
            //
            // EITHER name in EITHER position, by the one rule DataTreeUsage and the reviewer list
            // also skip by: a board they skip must never be publishable (code review, 2026-09-25).
            // ###########################################################################################
            if (SubmissionFileScopes.IsSharedFolderName(manifest.Manufacturer) ||
                SubmissionFileScopes.IsSharedFolderName(manifest.Hardware))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.reserved_folder",
                    manifest.SystemId,
                    $"A board cannot be named [{manifest.SystemId}]: that is where the shared files live."));
            }

            if ((manifest.BaseRevision?.Length ?? 0) > SubmissionFormat.MaximumRevisionLength ||
                (manifest.Rows?.RevisionDate?.Length ?? 0) > SubmissionFormat.MaximumRevisionLength)
            {
                findings.Add(SubmissionValidator.Error(
                    "revision.too_long",
                    string.Empty,
                    $"The board's revision date is longer than {SubmissionFormat.MaximumRevisionLength} characters."));
            }

            if ((manifest.Summary?.Length ?? 0) > SubmissionFormat.MaximumSummaryLength)
            {
                findings.Add(SubmissionValidator.Error(
                    "summary.too_long",
                    string.Empty,
                    $"The description of what changed is longer than {SubmissionFormat.MaximumSummaryLength} " +
                    "characters. Shorten it - the reviewer sees the full list of changes anyway."));
            }

            if (string.IsNullOrWhiteSpace(manifest.Hardware))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.hardware_missing", string.Empty, "The submission does not name the hardware."));
            }

            if (string.IsNullOrWhiteSpace(manifest.Board))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.board_missing", string.Empty, "The submission does not name the board."));
            }

            if (string.IsNullOrWhiteSpace(manifest.Summary))
            {
                // A warning, not an error: a reviewer can read the diff. But a submission that
                // says what it changed is reviewed far faster, so it is worth asking for.
                findings.Add(SubmissionValidator.Warning(
                    "summary.missing",
                    string.Empty,
                    "The submission does not say what was changed. A sentence helps whoever reviews it."));
            }
        }

        // -------------------------------------------------------------------------------------
        // Schematics.
        // -------------------------------------------------------------------------------------

        private static void ValidateSchematics(
            SubmissionManifest manifest,
            HashSet<string> files,
            Dictionary<string, string> filesIgnoringCase,
            List<ValidationFinding> findings)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardSchematicEntry schematic in manifest.Rows.Schematics)
            {
                if (string.IsNullOrWhiteSpace(schematic.SchematicName))
                {
                    findings.Add(SubmissionValidator.Error(
                        "schematic.unnamed", string.Empty, "A schematic row has no name."));

                    continue;
                }

                // Case-insensitive, because two schematics differing only by capitalisation would
                // be indistinguishable to a human reading the list - and highlights reference them
                // by name.
                if (!seen.Add(schematic.SchematicName))
                {
                    findings.Add(SubmissionValidator.Error(
                        "schematic.duplicate",
                        schematic.SchematicName,
                        $"More than one schematic is named [{schematic.SchematicName}]."));
                }

                SubmissionValidator.CheckFileReference(
                    schematic.SchematicImageFile,
                    $"schematic [{schematic.SchematicName}]",
                    files,
                    filesIgnoringCase,
                    findings,
                    required: true);
            }
        }

        // -------------------------------------------------------------------------------------
        // Components.
        // -------------------------------------------------------------------------------------

        private static void ValidateComponents(SubmissionManifest manifest, List<ValidationFinding> findings)
        {
            var seen = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (ComponentEntry component in manifest.Rows.Components)
            {
                if (string.IsNullOrWhiteSpace(component.BoardLabel))
                {
                    findings.Add(SubmissionValidator.Error(
                        "component.unlabelled",
                        string.Empty,
                        "A component row has no board label. The label is how everything else " +
                        "refers to it, so a row without one cannot be used."));

                    continue;
                }

                // ###########################################################################################
                // *** THE KEY IS LABEL + REGION, NOT LABEL ALONE (fixed 2026-09-23). ***
                //
                // A board label legitimately appears TWICE when the component differs between
                // regions: on the C64 250407, C70 is a ceramic capacitor on PAL and a film
                // capacitor on NTSC, and U19/Y1/R26/R52/R53 are region variants too. The
                // application has always supported this - ComponentListBuilder filters components
                // by region, so exactly one of the pair is ever shown - and the PUBLISHED board
                // ships with all six pairs in it.
                //
                // Keying on the label alone therefore rejected six rows of correct, already-
                // published data and made the whole board unsubmittable. Reported by the
                // maintainer on the first real submission of this board.
                //
                // A duplicate within ONE region is still an error: that is genuinely ambiguous,
                // because the region filter cannot tell the two rows apart. A blank region is its
                // own bucket - it means "all regions", so two blank-region rows collide as well.
                // ###########################################################################################
                string region = component.Region?.Trim() ?? string.Empty;

                // BoardDraftNaturalKeys.Separator, the same unit separator every other composite
                // key in this codebase uses, so a label or region containing an ordinary character
                // cannot fake a key boundary.
                string key = region.Length == 0
                    ? component.BoardLabel.Trim()
                    : component.BoardLabel.Trim() + BoardDraftNaturalKeys.Separator + region;

                if (seen.TryGetValue(key, out string? first))
                {
                    string where = region.Length == 0
                        ? "with no region"
                        : $"in region [{region}]";

                    findings.Add(SubmissionValidator.Error(
                        "component.duplicate_label",
                        component.BoardLabel,
                        $"More than one component is labelled [{component.BoardLabel}] {where} " +
                        $"(also [{first}]). Board labels must be unique within a region."));
                }
                else
                {
                    seen[key] = component.BoardLabel.Trim();
                }
            }
        }

        // -------------------------------------------------------------------------------------
        // Component images.
        // -------------------------------------------------------------------------------------

        private static void ValidateComponentImages(
            SubmissionManifest manifest,
            HashSet<string> files,
            Dictionary<string, string> filesIgnoringCase,
            List<ValidationFinding> findings)
        {
            var componentLabels = new HashSet<string>(
                manifest.Rows.Components.Select(component => component.BoardLabel),
                StringComparer.OrdinalIgnoreCase);

            foreach (ComponentImageEntry image in manifest.Rows.ComponentImages)
            {
                // ###########################################################################################
                // *** A COMPONENT-IMAGE ROW WITH NO FILE IS VALID - IT IS A NOTE (fixed 2026-09-23). ***
                //
                // This sheet carries a Note column as well as a File column, and a row may fill in
                // only the Note: "Compatible part-number: BZX55C2V7" on CR1/CR2, or Y1's warning
                // that probing the crystal directly can stall the machine. That is real repair
                // information with no picture attached.
                //
                // The application has always handled these - ComponentImageQueries has a dedicated
                // HasDisplayableImageFile predicate and simply skips such a row when rendering
                // images - and the published C64 250407 board ships with them. Requiring a file
                // here rejected correct, already-published data and blocked the whole submission.
                //
                // A row with NEITHER a file nor a note IS still reported: it declares nothing at
                // all, so it is a leftover rather than a note.
                // ###########################################################################################
                bool hasFile = !string.IsNullOrWhiteSpace(image.File);
                bool hasNote = !string.IsNullOrWhiteSpace(image.Note);

                if (hasFile || !hasNote)
                {
                    SubmissionValidator.CheckFileReference(
                        image.File,
                        $"component image for [{image.BoardLabel}]",
                        files,
                        filesIgnoringCase,
                        findings,
                        required: true);
                }

                // A warning rather than an error: an image for a component that is not in this
                // submission is usually a leftover, but it could be intentional during a staged
                // change, and the data still loads.
                if (!string.IsNullOrWhiteSpace(image.BoardLabel) &&
                    !componentLabels.Contains(image.BoardLabel))
                {
                    findings.Add(SubmissionValidator.Warning(
                        "image.orphan",
                        image.BoardLabel,
                        $"An image references component [{image.BoardLabel}], which is not in this " +
                        "submission. It will not be shown anywhere."));
                }
            }
        }

        // -------------------------------------------------------------------------------------
        // Highlights. The subtlest section, because the coordinates are STRINGS.
        // -------------------------------------------------------------------------------------

        private static void ValidateHighlights(SubmissionManifest manifest, List<ValidationFinding> findings)
        {
            var schematicNames = new HashSet<string>(
                manifest.Rows.Schematics.Select(schematic => schematic.SchematicName),
                StringComparer.OrdinalIgnoreCase);

            foreach (ComponentHighlightEntry highlight in manifest.Rows.ComponentHighlights)
            {
                string subject = $"{highlight.SchematicName} / {highlight.BoardLabel}";

                // A highlight naming a schematic that does not exist draws nothing, anywhere -
                // it is invisible rather than wrong-looking, which is why it has to be caught here
                // rather than noticed later.
                if (!schematicNames.Contains(highlight.SchematicName))
                {
                    findings.Add(SubmissionValidator.Error(
                        "highlight.unknown_schematic",
                        subject,
                        $"A highlight names schematic [{highlight.SchematicName}], which is not in " +
                        "this submission."));
                }

                if (!SubmissionValidator.TryParseCoordinate(highlight.X, out double x) ||
                    !SubmissionValidator.TryParseCoordinate(highlight.Y, out double y) ||
                    !SubmissionValidator.TryParseCoordinate(highlight.Width, out double width) ||
                    !SubmissionValidator.TryParseCoordinate(highlight.Height, out double height))
                {
                    findings.Add(SubmissionValidator.Error(
                        "highlight.unparseable",
                        subject,
                        $"A highlight has coordinates that cannot be read " +
                        $"(x=[{highlight.X}] y=[{highlight.Y}] w=[{highlight.Width}] h=[{highlight.Height}]). " +
                        "Numbers must use a dot as the decimal separator."));

                    continue;
                }

                // Zero-sized draws as nothing or as a hairline, and cannot be grabbed and moved -
                // the component looks like it has no highlight and there is no way to fix it from
                // the UI. The same defect WorklogDefaultAreaGeometry exists to prevent locally.
                if (width <= 0 || height <= 0)
                {
                    findings.Add(SubmissionValidator.Error(
                        "highlight.degenerate",
                        subject,
                        $"A highlight has no area (width {width}, height {height}). It would be " +
                        "invisible and could not be adjusted."));
                }

                if (x < 0 || y < 0)
                {
                    findings.Add(SubmissionValidator.Error(
                        "highlight.negative_origin",
                        subject,
                        $"A highlight starts outside the image (x {x}, y {y})."));
                }
            }
        }

        // -------------------------------------------------------------------------------------
        // Local files and links.
        // -------------------------------------------------------------------------------------

        private static void ValidateLocalFiles(
            SubmissionManifest manifest,
            HashSet<string> files,
            Dictionary<string, string> filesIgnoringCase,
            List<ValidationFinding> findings)
        {
            foreach (ComponentLocalFileEntry file in manifest.Rows.ComponentLocalFiles)
            {
                SubmissionValidator.CheckFileReference(
                    file.File, $"file for [{file.BoardLabel}]", files, filesIgnoringCase, findings, required: true);
            }

            foreach (BoardLocalFileEntry file in manifest.Rows.BoardLocalFiles)
            {
                SubmissionValidator.CheckFileReference(
                    file.File, "board file", files, filesIgnoringCase, findings, required: true);
            }
        }

        private static void ValidateLinks(SubmissionManifest manifest, List<ValidationFinding> findings)
        {
            foreach (ComponentLinkEntry link in manifest.Rows.ComponentLinks)
                SubmissionValidator.CheckLink(link.Url, $"link for [{link.BoardLabel}]", findings);

            // A board link is scoped by Category rather than by a component label - see
            // BoardLinkEntry.
            foreach (BoardLinkEntry link in manifest.Rows.BoardLinks)
                SubmissionValidator.CheckLink(link.Url, "board link", findings);
        }

        // ###########################################################################################
        // A link must be http or https.
        //
        // Not cosmetic: these are opened through ExternalTargetLauncher, which admits only
        // http/https/mailto and local paths inside the data root. A "file://" or "javascript:" URL
        // in contributed data is either a mistake or an attempt, and either way it will be refused
        // at click time with a warning nobody sees - so it is refused here where it can be
        // explained.
        // ###########################################################################################
        private static void CheckLink(string? link, string context, List<ValidationFinding> findings)
        {
            if (string.IsNullOrWhiteSpace(link))
                return;

            bool isWebLink = Uri.TryCreate(link, UriKind.Absolute, out Uri? uri)
                && (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));

            if (!isWebLink)
            {
                findings.Add(SubmissionValidator.Error(
                    "link.not_web",
                    link,
                    $"A {context} is [{link}], which is not an http:// or https:// address. " +
                    "CRT will not open anything else."));
            }
        }

        // -------------------------------------------------------------------------------------
        // Scale.
        // -------------------------------------------------------------------------------------

        private static void ValidateScale(SubmissionManifest manifest, List<ValidationFinding> findings)
        {
            if (manifest.Rows.Components.Count > SubmissionValidator.ImplausibleComponentCount)
            {
                findings.Add(SubmissionValidator.Warning(
                    "scale.components",
                    string.Empty,
                    $"The submission carries {manifest.Rows.Components.Count} components, far more than " +
                    "any board shipped with CRT. Worth a look before merging."));
            }

            if (manifest.Rows.Schematics.Count > SubmissionValidator.ImplausibleSchematicCount)
            {
                findings.Add(SubmissionValidator.Warning(
                    "scale.schematics",
                    string.Empty,
                    $"The submission carries {manifest.Rows.Schematics.Count} schematics, far more than " +
                    "any board shipped with CRT."));
            }

            if (manifest.Rows.Schematics.Count == 0 && manifest.Rows.Components.Count == 0)
            {
                findings.Add(SubmissionValidator.Error(
                    "submission.empty",
                    string.Empty,
                    "The submission contains no schematics and no components."));
            }
        }

        // -------------------------------------------------------------------------------------
        // Helpers.
        // -------------------------------------------------------------------------------------

        // ###########################################################################################
        // Checks that a referenced file is actually in the submission.
        //
        // THE CASE-MISMATCH BRANCH IS THE VALUABLE ONE. A reference that differs only in
        // capitalisation is the exact failure the strategy document warns about: it works on the
        // contributor's Windows machine and fails on the Linux server, after publication, on
        // somebody else's computer. Saying "the file is there but spelled differently" turns a
        // baffling bug into a one-character fix.
        // ###########################################################################################
        private static void CheckFileReference(
            string? reference,
            string context,
            HashSet<string> files,
            Dictionary<string, string> filesIgnoringCase,
            List<ValidationFinding> findings,
            bool required)
        {
            if (string.IsNullOrWhiteSpace(reference))
            {
                if (required)
                {
                    findings.Add(SubmissionValidator.Error(
                        "file.unreferenced", context, $"The {context} does not name a file."));
                }

                return;
            }

            if (files.Contains(reference))
                return;

            if (filesIgnoringCase.TryGetValue(reference, out string? actual))
            {
                findings.Add(SubmissionValidator.Error(
                    "file.case_mismatch",
                    reference,
                    $"The {context} names [{reference}], but the file in the submission is spelled " +
                    $"[{actual}]. Capitalisation matters on the server even though it does not on " +
                    "Windows, so this would work for you and fail for everyone else."));

                return;
            }

            findings.Add(SubmissionValidator.Error(
                "file.missing",
                reference,
                $"The {context} names [{reference}], which is not in the submission."));
        }

        // ###########################################################################################
        // Parses a coordinate written as text.
        //
        // INVARIANT CULTURE ONLY. These values are written by a Danish, German or French
        // contributor's machine as readily as an English one, and "1,5" parsed under a comma-
        // decimal culture is 1.5 while under an invariant one it is 15 - a highlight ten times too
        // wide, with nothing failing. The data format is invariant, so a comma is a malformed
        // value and must be reported rather than reinterpreted.
        // ###########################################################################################
        private static bool TryParseCoordinate(string? value, out double result)
        {
            result = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            return double.TryParse(
                value, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        // An empty part is canonical here - "missing" is its own finding, and reporting it twice
        // under two codes helps nobody.
        private static bool IsCanonical(string? part) =>
            string.IsNullOrEmpty(part) ||
            string.Equals(part, NewSystemIdentity.SanitizePathSegment(part), StringComparison.Ordinal);

        private static ValidationFinding Error(string code, string subject, string message) =>
            new() { Severity = ValidationSeverity.Error, Code = code, Subject = subject, Message = message };

        private static ValidationFinding Warning(string code, string subject, string message) =>
            new() { Severity = ValidationSeverity.Warning, Code = code, Subject = subject, Message = message };
    }
}
