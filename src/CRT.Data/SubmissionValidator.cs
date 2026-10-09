using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Checks a submission's ROWS before any human sees it. Pure: rows and a set of known file
    // paths in, findings out. It opens nothing and reads no images, so every rule is a unit test.
    //
    // *** THIS IS THE HIGHEST-LEVERAGE WORK IN THE WHOLE PLAN. *** NewContributeStrategy.md says
    // so outright, and the reason is arithmetic: the project owner is one volunteer. Every submission
    // rejected automatically with a clear explanation is maintainer time not spent, and every bad
    // submission that reaches a human costs far more than the contributor saved by not checking.
    // A rejection here is also FASTER for the contributor than a review queue.
    //
    // HARD FAILURES ARE REJECTED AUTOMATICALLY AND NEVER QUEUED. Warnings are attached for the
    // maintainer to weigh. The line between them is "could a reasonable contributor have meant
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

            SubmissionValidator.ValidateIdentity(manifest, findings);

            // ###########################################################################################
            // *** THE ROW RULES ARE BoardDataChecks', IN ITS SUBMISSION SCOPE (2026-10-02). *** The
            // Drafts tab's table now shows the same rules on the cells they are about, so they were
            // moved there rather than copied: one set, so an error the contributor sees in the table
            // is exactly one this refuses. The scope is these rules exactly - codes, subjects,
            // messages and order unchanged - which every test of this class still holds.
            //
            // The supplied paths compare ORDINALLY, with a case-insensitive second look only for
            // the better message (SuppliedFileLookup) - see the header for why.
            // ###########################################################################################
            foreach (BoardDataProblem problem in BoardDataChecks.Check(
                         BoardCheckRows.From(manifest.Rows),
                         new SuppliedFileLookup(suppliedPaths),
                         BoardCheckScope.Submission))
            {
                findings.Add(new ValidationFinding
                {
                    Severity = problem.Level == BoardProblemLevel.Error ? ValidationSeverity.Error : ValidationSeverity.Warning,
                    Code = problem.Code,
                    Subject = problem.Subject,
                    Message = problem.Message
                });
            }

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
            // THE BoardId IS REQUIRED AND MUST BE THE RIGHT SHAPE.
            //
            // It is "Manufacturer/Hardware/Board" and the client can always compute it from what
            // the contributor typed, so an absent or malformed one is a broken client rather than
            // a new board - a NEW board carries an id too, and is new because no `boards` row
            // holds that id yet.
            //
            // This matters more than an ordinary field check: the value is a DATABASE PRIMARY KEY
            // (and a foreign key from `submissions`), and it is used to resolve a folder in the
            // published tree. A malformed one is either a broken client or an attempt to attach a
            // submission to something it does not belong to.
            //
            // Shape is all that can be decided here. Whether the id names a board that EXISTS,
            // and whether this contributor may submit to it, are database questions the server
            // answers separately.
            // ###########################################################################################
            if (!BoardDescriptorRules.IsValidBoardId(manifest.BoardId))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.board_id_malformed",
                    string.Empty,
                    "The submission does not carry a valid board identifier. " +
                    "This is a fault in the submitting application rather than in your data."));
            }

            // The id is built from these three, so a disagreement means the client assembled the
            // manifest inconsistently - and the id is what everything keys off, so the parts would
            // silently be the wrong ones on any screen that showed them.
            if (BoardDescriptorRules.IsValidBoardId(manifest.BoardId)
                && !string.Equals(
                    manifest.BoardId,
                    BoardDescriptorRules.BuildBoardId(manifest.Manufacturer, manifest.Hardware, manifest.Board),
                    StringComparison.Ordinal))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.board_id_mismatch",
                    string.Empty,
                    "The submission's board identifier does not match the manufacturer, hardware " +
                    "and board it names. This is a fault in the submitting application."));
            }

            // ###########################################################################################
            // *** THE PARTS MUST ALREADY BE IN THEIR CANONICAL FORM (security review, 2026-09-25). ***
            //
            // BuildBoardId trims and collapses whitespace before joining, so "Commodore" followed
            // by a thousand spaces produced a VALID, MATCHING id - and the raw thousand-character
            // part then went into boards.manufacturer, a VARCHAR(100), failing the insert with a
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
            // "Generic shared files", or hardware called "Shared files", would make this board's
            // own folder the same place as a shared one - and a submission may change its own folder
            // freely, so that would turn the shared-folder rules in SubmissionFileScope inside out.
            //
            // EITHER name in EITHER position, by the one rule DataTreeUsage and the maintainer list
            // also skip by: a board they skip must never be publishable (code review, 2026-09-25).
            // ###########################################################################################
            if (SubmissionFileScopes.IsSharedFolderName(manifest.Manufacturer) ||
                SubmissionFileScopes.IsSharedFolderName(manifest.Hardware))
            {
                findings.Add(SubmissionValidator.Error(
                    "identity.reserved_folder",
                    manifest.BoardId,
                    $"A board cannot be named [{manifest.BoardId}]: that is where the shared files live."));
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
                    "characters. Shorten it - the maintainer sees the full list of changes anyway."));
            }

            // A new board's notes go into the main Excel data file's notes column, which holds no
            // more than this (MasterListing.IsWritableRow) - refused here, where the contributor
            // can shorten them, rather than when the maintainer tries to save the placement.
            if ((manifest.HardwareNotes?.Length ?? 0) > MasterListing.MaximumNotesLength)
            {
                findings.Add(SubmissionValidator.Error(
                    "notes.too_long",
                    string.Empty,
                    $"The hardware notes are longer than {MasterListing.MaximumNotesLength} characters. Shorten them."));
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
                // A warning, not an error: a maintainer can read the diff. But a submission that
                // says what it changed is reviewed far faster, so it is worth asking for.
                findings.Add(SubmissionValidator.Warning(
                    "summary.missing",
                    string.Empty,
                    "The submission does not say what was changed. A sentence helps whoever reviews it."));
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

        // An empty part is canonical here - "missing" is its own finding, and reporting it twice
        // under two codes helps nobody.
        private static bool IsCanonical(string? part) =>
            string.IsNullOrEmpty(part) ||
            string.Equals(part, NewBoardIdentity.SanitizePathSegment(part), StringComparison.Ordinal);

        private static ValidationFinding Error(string code, string subject, string message) =>
            new() { Severity = ValidationSeverity.Error, Code = code, Subject = subject, Message = message };

        private static ValidationFinding Warning(string code, string subject, string message) =>
            new() { Severity = ValidationSeverity.Warning, Code = code, Subject = subject, Message = message };
    }
}
