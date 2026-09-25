using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH FILES a submission may carry and WHERE (security review, 2026-09-25). Pure: a manifest
    // and a view of the published tree in, findings out.
    //
    // SubmissionPathRules answers "is this path SHAPED safely" - relative, contained, no
    // traversal. That turned out not to be enough, because a perfectly shaped path can still:
    //
    //   1. name ANOTHER BOARD'S file ("Commodore/C64/250407/..." inside a submission to the
    //      Amstrad CPC), which the publish then overwrote - see SubmissionFileScope;
    //   2. be a file NO ROW USES, which the review screen had no way to show, so a maintainer
    //      approved it without ever seeing it;
    //   3. be any type at all - a dot-file such as ".htaccess", which the web server hosting the
    //      data tree honours, or an executable that syncs to every user's disk;
    //   4. differ from a published path only by capitalisation, which Linux keeps apart and every
    //      Windows and macOS client merges.
    //
    // Each is refused here, at create (before anything is uploaded) and again by PublishPlan (the
    // last gate before the tree is written). The same rules in both places, from this one class.
    //
    // *** EVERY RULE WAS CHECKED AGAINST EVERY SHIPPED BOARD. *** This project has twice shipped a
    // rule that rejected correct, already-published data (region-variant labels, note-only image
    // rows). SubmissionFileRulesShippedDataTests reads every board in Assets/Data and asserts none
    // of them is refused - which is how the C128DCR's cross-board scope-baseline references were
    // found, and why an UNCHANGED foreign file is allowed.
    // ###########################################################################################
    public static class SubmissionFileRules
    {
        // ###########################################################################################
        // The file types a submission may carry. An ALLOWLIST, never a blocklist: anything not
        // named here - including a file with no extension at all - is refused.
        //
        // Derived from what board rows actually reference: every shipped board, read in full,
        // cites only .png, .jpg, .jpeg, .gif, .pdf, .txt and one shared .html page. .bmp and .webp
        // are images the review screen and the app already draw. KiCad projects, scope captures
        // and the MiniPro catalogue are found by scanning folders rather than through rows, so they
        // never travel in a submission and are deliberately absent.
        //
        // *** NOT .svg, .json, .xml or .xlsx. *** SVG can carry script. A .json is the board's own
        // highlight sidecar or its system.json, both generated at publish and never uploaded - an
        // uploaded one would rewrite a board's highlights or its maintainer list. A workbook is
        // generated from the rows (PublishPlan refuses one with its own message).
        //
        // SubmissionContentRules must know the signature of every entry here - see its default arm.
        // ###########################################################################################
        public static readonly IReadOnlySet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp",
            ".pdf", ".txt", ".html", ".htm"
        };

        // ###########################################################################################
        // Is this path's NAME acceptable - allowed type, and nothing hidden anywhere along it?
        //
        // *** A DOT ANYWHERE AT THE START OF A SEGMENT IS REFUSED, not just on the file. *** ".git/x"
        // and ".well-known/y" are folders the web server and tooling treat specially, and a
        // dot-file such as ".htaccess" is read by Apache as configuration for the folder it sits
        // in. The checksum manifest already skips every dot-path, so nothing legitimate is lost.
        // ###########################################################################################
        public static bool TryCheckName(string? path, out string code, out string reason)
        {
            code = string.Empty;
            reason = string.Empty;

            if (string.IsNullOrWhiteSpace(path))
            {
                code = "file.no_name";
                reason = "A file in the submission has no name.";
                return false;
            }

            foreach (string segment in path.Split('/'))
            {
                if (segment.StartsWith('.'))
                {
                    code = "file.hidden_name";
                    reason = $"[{path}] has a name starting with a dot. Hidden files cannot be submitted.";
                    return false;
                }
            }

            string extension = Path.GetExtension(path);

            if (extension.Length == 0)
            {
                code = "file.no_type";
                reason = $"[{path}] has no file type (no extension such as .png or .pdf), so it cannot be submitted.";
                return false;
            }

            if (!SubmissionFileRules.AllowedExtensions.Contains(extension))
            {
                code = "file.type_not_allowed";
                reason =
                    $"[{path}] is a {extension} file, which cannot be submitted. Board data may carry " +
                    $"these types: {string.Join(", ", SubmissionFileRules.AllowedExtensions.Order(StringComparer.Ordinal))}.";
                return false;
            }

            return true;
        }

        // ###########################################################################################
        // The files a manifest's ROWS name - the only files it may carry.
        //
        // Through the SHARED collector, the same one the client builds its file list from, so a
        // client that sends exactly what its rows reference is never refused by this rule.
        // ###########################################################################################
        public static IReadOnlySet<string> ReferencedFiles(SubmissionManifest manifest)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return new HashSet<string>(
                SubmissionManifestBuilder.CollectReferencedFiles(PublishMerge.Build(manifest, published: null)),
                StringComparer.Ordinal);
        }

        // ###########################################################################################
        // Every problem with the manifest's FILES, all at once.
        //
        // `tree` is what is published now. Null means it could not be consulted, and then every
        // foreign file is refused - the one rule that needs the tree fails CLOSED without it - and
        // case variants go unchecked here (PublishPlan checks them again against the real tree).
        //
        // Paths that are not even SHAPED safely are skipped (SubmissionPathRules.IsSafelyShaped):
        // SubmissionPathRules.ValidateManifestPaths reports those, and every caller runs it first. A
        // second finding about the same bad path under another code only contradicted the first.
        // ###########################################################################################
        public static IReadOnlyList<ValidationFinding> ValidateManifestFiles(
            SubmissionManifest manifest,
            PublishedTreeView? tree)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            var findings = new List<ValidationFinding>();
            IReadOnlySet<string> referenced = SubmissionFileRules.ReferencedFiles(manifest);
            var seen = new HashSet<string>(StringComparer.Ordinal);

            foreach (SubmissionFile file in manifest.Files)
            {
                if (string.IsNullOrWhiteSpace(file.Path) || !seen.Add(file.Path))
                    continue;

                if (!SubmissionPathRules.IsSafelyShaped(file.Path, out _))
                    continue;

                if (!SubmissionFileRules.TryCheckName(file.Path, out string nameCode, out string nameReason))
                {
                    findings.Add(SubmissionFileRules.Error(nameCode, file.Path, nameReason));
                    continue;
                }

                if (!referenced.Contains(file.Path))
                {
                    findings.Add(SubmissionFileRules.Error(
                        "file.not_used",
                        file.Path,
                        $"[{file.Path}] is in the submission, but nothing in the board uses it. Only files " +
                        "the board's rows name can be submitted."));

                    continue;
                }

                if (SubmissionFileScopes.Classify(manifest, file.Path) == SubmissionFileScope.Foreign)
                {
                    SubmissionFileRules.CheckForeign(file, tree, findings);
                    continue;
                }

                string? variant = tree?.FindCaseVariant(file.Path);

                if (variant is not null)
                    findings.Add(SubmissionFileRules.CaseVariant(file.Path, variant));
            }

            // The system's own folder too, which a submission carrying no files at all would
            // otherwise never check - "commodore/C64/250407" beside a published "Commodore/...".
            string systemFolder = $"{manifest.Manufacturer}/{manifest.Hardware}/{manifest.Board}";
            string? folderVariant = tree?.FindCaseVariant(systemFolder);

            if (folderVariant is not null)
            {
                findings.Add(SubmissionFileRules.Error(
                    "identity.case_collision",
                    systemFolder,
                    $"The board [{systemFolder}] differs only by capitalisation from the published " +
                    $"[{folderVariant}]. On Windows and macOS the two would be the same folder, so it " +
                    "cannot be submitted as a separate board. Use the published spelling."));
            }

            return findings;
        }

        // ###########################################################################################
        // A file outside this board and the shared folders: allowed ONLY when it is byte-identical
        // to what is published there. Citing another board's file is how real data is shaped;
        // CHANGING it from here is what must never happen. See SubmissionFileScope's header.
        // ###########################################################################################
        private static void CheckForeign(SubmissionFile file, PublishedTreeView? tree, List<ValidationFinding> findings)
        {
            string? published = tree?.HashOf(file.Path);

            if (published is null)
            {
                findings.Add(SubmissionFileRules.Error(
                    "file.other_board",
                    file.Path,
                    $"[{file.Path}] belongs to another board. A submission can only add or change files " +
                    "in its own board's folder and in the shared folders."));

                return;
            }

            if (!string.Equals(published, file.Sha256, StringComparison.Ordinal))
            {
                findings.Add(SubmissionFileRules.Error(
                    "file.other_board_changed",
                    file.Path,
                    $"[{file.Path}] belongs to another board and differs from the published copy. Another " +
                    "board's files can be used as they are, but not changed from here."));
            }
        }

        internal static ValidationFinding CaseVariant(string path, string variant) =>
            SubmissionFileRules.Error(
                "path.case_collision_published",
                path,
                $"[{path}] differs only by capitalisation from the published [{variant}]. On Windows " +
                "and macOS the two are the same file, so this would overwrite it for most users. Use " +
                "the published spelling.");

        private static ValidationFinding Error(string code, string subject, string message) =>
            new() { Severity = ValidationSeverity.Error, Code = code, Subject = subject, Message = message };
    }
}
