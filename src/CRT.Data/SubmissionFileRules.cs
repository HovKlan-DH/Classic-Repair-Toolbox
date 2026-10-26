using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH FILES a submission may carry (security review, 2026-09-25). Pure: a manifest and a view
    // of the published board in, findings out.
    //
    // *** IN THE CRT 4.0 TREE A SUBMITTED FILE IS NAMED BY ITS BOARD (2026-10-25). *** A file's
    // `Path` is the board's own name for it ("Component images/7406.png"), the name its rows cite;
    // its bytes are stored under their checksum (DataTreeLayout.StoredFilePath), so a submission
    // never names a place in the tree and can never touch another board's file. The rules that
    // stopped a path reaching another board or a shared folder (SubmissionFileScope's Foreign and
    // shared files) went with the shared folders - every board owns its files (owner decision,
    // 2026-10-24).
    //
    // SubmissionPathRules answers "is this name SHAPED safely" - relative, no traversal, no
    // control characters, no two names that differ only by capitalisation. A safely shaped name can
    // still:
    //
    //   1. be a file NO ROW USES, which the review screen had no way to show, so a maintainer
    //      approved it without ever seeing it;
    //   2. be any type at all - a dot-file such as ".htaccess", or an executable that syncs to
    //      every user's disk (a name is also where CRT puts the file in a draft or a KiCad copy);
    //   3. give a NEW board a folder that differs from a published one only by capitalisation,
    //      which Linux keeps apart and every Windows and macOS client merges.
    //
    // Each is refused here, at create (before anything is uploaded) and again by PublishPlan (the
    // last gate before the tree is written). The same rules in both places, from this one class.
    // ###########################################################################################
    public static class SubmissionFileRules
    {
        // ###########################################################################################
        // The file types a submission may carry. An ALLOWLIST, never a blocklist: anything not
        // named here - including a file with no extension at all - is refused.
        //
        // Derived from what board rows actually reference: every shipped board, read in full,
        // cites only .png, .jpg, .jpeg, .gif, .pdf, .txt and one shared .html page. .bmp and .webp
        // are images the review screen and the app already draw. The MiniPro catalogue is found by
        // scanning folders rather than through rows, so it never travels in a submission. *** The
        // board's own KiCad DATA travels since 2026-09-26, but NOT through this list: the KiCad
        // types CRT reads are admitted only in the board's "KiCad data" folder
        // (SubmissionKiCadFiles.IsSubmittable), never anywhere a row could cite one. ***
        //
        // *** NOT .svg, .json, .xml or .xlsx. *** SVG can carry script. The board file and the main
        // file are .json, written by the publish from the rows and never uploaded. A workbook is
        // what older CRTs read, and no 4.0 board names one.
        //
        // SubmissionContentRules must know the signature of every entry here - see its default arm.
        // ###########################################################################################
        public static readonly IReadOnlySet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp",
            ".pdf", ".txt", ".html", ".htm"
        };

        // ###########################################################################################
        // Is this name acceptable - allowed type, and nothing hidden anywhere along it?
        //
        // *** A DOT ANYWHERE AT THE START OF A SEGMENT IS REFUSED, not just on the file. *** ".git/x"
        // and ".well-known/y" are folders tooling treats specially, a dot-file such as ".htaccess"
        // is read by Apache as configuration for the folder it sits in, and CRT's own bookkeeping in
        // a draft is ".crt-*". The checksum manifest skips every dot-path, so nothing legitimate is
        // lost.
        //
        // `kiCadProjectFile` widens the allowlist - passed true ONLY for a file
        // SubmissionKiCadFiles.IsSubmittable already said yes to (2026-09-26).
        // ###########################################################################################
        public static bool TryCheckName(string? path, out string code, out string reason, bool kiCadProjectFile = false)
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

            if (kiCadProjectFile)
                return true;

            if (!SubmissionFileRules.AllowedExtensions.Contains(extension))
            {
                code = "file.type_not_allowed";
                reason =
                    $"[{path}] is a {extension} file, which cannot be submitted. Board data may carry " +
                    $"these types: {string.Join(", ", SubmissionFileRules.AllowedExtensions.Order(StringComparer.Ordinal))}, " +
                    "plus KiCad project files in the board's own \"KiCad data\" folder.";
                return false;
            }

            return true;
        }

        // ###########################################################################################
        // The files a manifest's ROWS name - the only files besides its KiCad data it may carry.
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
        // Is this one of the manifest's files the submission may carry: a KiCad file by
        // SubmissionKiCadFiles' rule (against the published board's file of that name), or a file a
        // row cites - of an allowed type and no hidden name either way? The reason when it is not.
        // ###########################################################################################
        public static bool TryAdmit(
            SubmissionManifest manifest,
            SubmissionFile file,
            IReadOnlySet<string> referenced,
            PublishedTreeView? board,
            out string code,
            out string reason)
        {
            ArgumentNullException.ThrowIfNull(manifest);
            ArgumentNullException.ThrowIfNull(file);
            ArgumentNullException.ThrowIfNull(referenced);

            bool kiCadProjectFile = SubmissionKiCadFiles.IsSubmittable(manifest, file.Path, board?.HashOf(file.Path), file.Sha256);

            if (!SubmissionFileRules.TryCheckName(file.Path, out code, out reason, kiCadProjectFile))
                return false;

            if (!kiCadProjectFile && !referenced.Contains(file.Path))
            {
                code = "file.not_used";
                reason =
                    $"[{file.Path}] is in the submission, but nothing in the board uses it. Only files " +
                    "the board's rows name, and its KiCad data, can be submitted.";

                return false;
            }

            return true;
        }

        // ###########################################################################################
        // Every problem with the manifest's FILES, all at once.
        //
        // `board` is the published board's files by name (PublishedTreeView), and the published
        // tree's folders - for the board folder's capitalisation. Null means it could not be
        // consulted: no file of the KiCad folder is then admitted on the strength of being published
        // already, and the folder goes unchecked here (PublishPlan checks it again).
        //
        // Names that are not even SHAPED safely are skipped (SubmissionPathRules.IsSafelyShaped):
        // SubmissionPathRules.ValidateManifestPaths reports those, and every caller runs it first. A
        // second finding about the same bad name under another code only contradicted the first.
        // ###########################################################################################
        public static IReadOnlyList<ValidationFinding> ValidateManifestFiles(
            SubmissionManifest manifest,
            PublishedTreeView? board)
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

                if (!SubmissionFileRules.TryAdmit(manifest, file, referenced, board, out string code, out string reason))
                    findings.Add(SubmissionFileRules.Error(code, file.Path, reason));
            }

            // The board's own folder, which a NEW board would create - "commodore/C64/250407" beside
            // a published "Commodore/...".
            string boardFolder = $"{manifest.Manufacturer}/{manifest.Hardware}/{manifest.Board}";
            string? folderVariant = board?.FindCaseVariant(boardFolder);

            if (folderVariant is not null)
            {
                findings.Add(SubmissionFileRules.Error(
                    "identity.case_collision",
                    boardFolder,
                    $"The board [{boardFolder}] differs only by capitalisation from the published " +
                    $"[{folderVariant}]. On Windows and macOS the two would be the same folder, so it " +
                    "cannot be submitted as a separate board. Use the published spelling."));
            }

            return findings;
        }

        private static ValidationFinding Error(string code, string subject, string message) =>
            new() { Severity = ValidationSeverity.Error, Code = code, Subject = subject, Message = message };
    }
}
