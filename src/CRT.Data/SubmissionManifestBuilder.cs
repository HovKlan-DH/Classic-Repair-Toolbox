using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Turns a drafted system into the SubmissionManifest the server expects. Pure: it takes the
    // merged board data, the set of files with their hashes already computed, and the identity -
    // and returns the manifest. It reads no files and computes no hashes, so every rule is a unit
    // test.
    //
    // WHY HASHING IS NOT DONE HERE. Hashing a 76 MB system is seconds of work that must happen off
    // the UI thread with progress reported, and it touches the filesystem. Keeping it outside
    // leaves this class testable without fixtures and leaves the caller free to report progress
    // however it likes - the same split ExportOverlayGeometry uses against the PDF exporter.
    //
    // THE MANIFEST IS THE COMPLETE INTENDED STATE, not a list of changes. That is the contract's
    // central idea: the server diffs the base revision against what it is told the system should
    // be, so the client never computes a diff at all. It also means a file the draft did NOT touch
    // still has to appear here - leaving it out would read as a deletion.
    //
    // WHICH FILES ARE REFERENCED IS DERIVED FROM THE ROWS, never from a directory walk. A walk
    // would sweep up editor backups, thumbnails caches and whatever else happens to sit in the
    // folder, and would upload a contributor's unrelated files to a public server. Only what the
    // data actually names is sent.
    // ###########################################################################################
    public static class SubmissionManifestBuilder
    {
        // ###########################################################################################
        // Builds the manifest.
        //
        // merged is the board data as it should read after the change - normally
        // BoardDraftApplier.ApplyDraft(official, draft).
        //
        // fileHashes maps each referenced relative path to its SHA-256 and size. The caller
        // computes these; a path the rows reference but this dictionary lacks becomes a finding
        // rather than a silent omission, because a manifest naming a file the server will never
        // receive fails at finalise with a far less helpful message.
        // ###########################################################################################
        public static SubmissionManifestBuildResult Build(
            BoardData merged,
            SubmissionIdentity identity,
            IReadOnlyDictionary<string, SubmissionFileInfo> fileHashes,
            IReadOnlyList<SubmissionRename>? renames = null,

            // ###########################################################################################
            // The draft's KiCad calibrations, which CANNOT come from `merged`.
            //
            // BoardData has no calibration section - they live only in the draft and, once
            // published, only in the JSON sidecar. So they arrive alongside the board rather than
            // inside it. OPTIONAL and TRAILING deliberately: every existing caller keeps working
            // unchanged, which is what made closing this gap safe to do without touching them.
            //
            // KiCadCalibrationDraftWriter.CollectForSubmission is what a caller passes here.
            // ###########################################################################################
            IReadOnlyList<KiCadCalibrationEntry>? calibrations = null)
        {
            ArgumentNullException.ThrowIfNull(merged);
            ArgumentNullException.ThrowIfNull(identity);
            ArgumentNullException.ThrowIfNull(fileHashes);

            var problems = new List<string>();

            IReadOnlyList<string> referenced = SubmissionManifestBuilder.CollectReferencedFiles(merged);

            var files = new List<SubmissionFile>();

            foreach (string path in referenced)
            {
                if (!fileHashes.TryGetValue(path, out SubmissionFileInfo info))
                {
                    // Named rather than dropped. A contributor whose schematic image is missing
                    // needs to know which one, not "the submission failed".
                    problems.Add(
                        $"The data references [{path}], but that file could not be found. " +
                        "It may have been moved or deleted since it was added.");

                    continue;
                }

                files.Add(new SubmissionFile
                {
                    Path = path,
                    Sha256 = info.Sha256,
                    SizeBytes = info.SizeBytes
                });
            }

            var manifest = new SubmissionManifest
            {
                FormatVersion = SubmissionFormat.CurrentVersion,
                SystemId = identity.SystemId,
                Manufacturer = identity.Manufacturer,
                Hardware = identity.Hardware,
                Board = identity.Board,
                BaseRevision = identity.BaseRevision,
                Summary = identity.Summary,
                ApplicationVersion = identity.ApplicationVersion,
                CreatedUtc = identity.CreatedUtc,
                Files = files,
                Renames = renames?.ToList() ?? [],
                Rows = new SubmissionRows
                {
                    // The board's own revision date travels with the rows, so a contributor who
                    // corrects it can actually publish that correction. Added 2026-09-22; before
                    // that the field did not exist and a revision-date change was both invisible
                    // to the maintainer and unpublishable.
                    //
                    // Calibrations come from the CALLER, not from `merged` - see the parameter's
                    // own comment for why a BoardData cannot carry them.
                    RevisionDate = merged.RevisionDate ?? string.Empty,
                    KiCadCalibrations = calibrations?.ToList() ?? [],

                    Schematics = merged.Schematics.ToList(),
                    Components = merged.Components.ToList(),
                    ComponentImages = merged.ComponentImages.ToList(),
                    ComponentHighlights = merged.ComponentHighlights.ToList(),
                    ComponentLocalFiles = merged.ComponentLocalFiles.ToList(),
                    ComponentLinks = merged.ComponentLinks.ToList(),
                    BoardLocalFiles = merged.BoardLocalFiles.ToList(),
                    BoardLinks = merged.BoardLinks.ToList(),
                    Credits = merged.Credits.ToList(),
                    KiCadImportantSignals = merged.KiCadImportantSignals.ToList()
                }
            };

            return new SubmissionManifestBuildResult(manifest, problems);
        }

        // ###########################################################################################
        // Every file path the board data references, de-duplicated, in a stable order.
        //
        // CASE-SENSITIVE de-duplication, like everything else touching paths: "U8.png" and "u8.PNG"
        // are two files to the server, and collapsing them here would silently drop one and leave
        // the row that named it pointing at nothing.
        //
        // ORDERED, so the same system produces the same manifest twice. That matters for a
        // contributor comparing two submissions, and it makes the hash-negotiation step
        // reproducible when something goes wrong and has to be diagnosed.
        // ###########################################################################################
        public static IReadOnlyList<string> CollectReferencedFiles(BoardData merged)
        {
            ArgumentNullException.ThrowIfNull(merged);

            var paths = new SortedSet<string>(StringComparer.Ordinal);

            foreach (BoardSchematicEntry schematic in merged.Schematics)
                SubmissionManifestBuilder.AddIfPresent(paths, schematic.SchematicImageFile);

            foreach (ComponentImageEntry image in merged.ComponentImages)
                SubmissionManifestBuilder.AddIfPresent(paths, image.File);

            foreach (ComponentLocalFileEntry file in merged.ComponentLocalFiles)
                SubmissionManifestBuilder.AddIfPresent(paths, file.File);

            foreach (BoardLocalFileEntry file in merged.BoardLocalFiles)
                SubmissionManifestBuilder.AddIfPresent(paths, file.File);

            return paths.ToList();
        }

        // ###########################################################################################
        // A human-readable summary of what is about to be sent, for the confirmation dialog.
        //
        // The contributor is about to publish their work to a server other people will download
        // from, so "exactly what will be sent" is a requirement in NewContributeStrategy.md rather
        // than a nicety. The counts here are what that dialog shows.
        // ###########################################################################################
        public static SubmissionPreview Preview(
            SubmissionManifest manifest,
            IReadOnlyCollection<string> hashesTheServerAlreadyHas)
        {
            ArgumentNullException.ThrowIfNull(manifest);
            ArgumentNullException.ThrowIfNull(hashesTheServerAlreadyHas);

            var held = new HashSet<string>(hashesTheServerAlreadyHas, StringComparer.Ordinal);

            // De-duplicated by hash: one blob referenced from three rows uploads once, and a
            // preview that counted it three times would overstate the upload.
            var distinct = new Dictionary<string, long>(StringComparer.Ordinal);

            foreach (SubmissionFile file in manifest.Files)
                distinct.TryAdd(file.Sha256, file.SizeBytes);

            long toUpload = distinct
                .Where(pair => !held.Contains(pair.Key))
                .Sum(pair => pair.Value);

            int newCount = distinct.Count(pair => !held.Contains(pair.Key));

            return new SubmissionPreview(
                manifest.Rows.Schematics.Count,
                manifest.Rows.Components.Count,
                manifest.Files.Count,
                newCount,
                distinct.Count - newCount,
                toUpload);
        }

        // ###########################################################################################
        // Formats a byte count for the confirmation dialog.
        //
        // Invariant culture with one decimal: this is shown to a person, and a locale that uses a
        // comma decimal separator would be fine - but the app has no localisation, so mixing
        // conventions between machines would be the only inconsistency on screen.
        // ###########################################################################################
        public static string FormatSize(long bytes)
        {
            if (bytes < 1024)
                return $"{bytes} bytes";

            if (bytes < 1024 * 1024)
                return string.Create(CultureInfo.InvariantCulture, $"{bytes / 1024.0:F1} KB");

            if (bytes < 1024L * 1024 * 1024)
                return string.Create(CultureInfo.InvariantCulture, $"{bytes / (1024.0 * 1024):F1} MB");

            return string.Create(CultureInfo.InvariantCulture, $"{bytes / (1024.0 * 1024 * 1024):F2} GB");
        }

        private static void AddIfPresent(SortedSet<string> paths, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
                paths.Add(value.Trim());
        }
    }

    // ###########################################################################################
    // Who and what is being submitted. Separate from the manifest so the caller assembles it from
    // wherever these values live (DataManager, the draft's registration, the csproj version)
    // without this class knowing about any of them.
    // ###########################################################################################
    public sealed class SubmissionIdentity
    {
        public string SystemId { get; init; } = string.Empty;
        public string Manufacturer { get; init; } = string.Empty;
        public string Hardware { get; init; } = string.Empty;
        public string Board { get; init; } = string.Empty;
        public string BaseRevision { get; init; } = string.Empty;
        public string Summary { get; init; } = string.Empty;
        public string ApplicationVersion { get; init; } = string.Empty;
        public DateTimeOffset CreatedUtc { get; init; }
    }

    public readonly record struct SubmissionFileInfo(string Sha256, long SizeBytes);

    // ###########################################################################################
    // The manifest, plus anything that stopped it being complete. Problems are NOT findings:
    // findings come from the server's validator, while these are local failures - a file the rows
    // name that is not on this machine any more.
    // ###########################################################################################
    public sealed record SubmissionManifestBuildResult(
        SubmissionManifest Manifest,
        IReadOnlyList<string> Problems)
    {
        public bool IsComplete => this.Problems.Count == 0;
    }

    // ###########################################################################################
    // Turns a manifest's rows back into a BoardData - the inverse of what SubmissionManifestBuilder
    // does, needed by anything that has to treat a submission AS a board.
    //
    // Two things need it and neither is the client: the review screen compares the submitted board
    // against the published one (Phase 5 task 3), and publishing writes the submitted board into
    // the data tree (task 6). Both would otherwise re-implement this mapping, and the two copies
    // would drift the first time a section was added - the exact defect SubmissionContract.cs
    // exists to prevent, one layer down.
    //
    // *** THE REVISION DATE IS NOT IN THE ROWS, AND IS PASSED IN. *** SubmissionRows carries the
    // ten board sections and no revision date; the manifest's own BaseRevision is the date the
    // contributor STARTED from, which is a different thing and must not be mistaken for it -
    // writing the base revision back out would silently revert a revision-date change and make
    // the board claim to be older than it is. The caller supplies what it means, and a blank is a
    // legitimate answer.
    //
    // KiCadCalibrations are deliberately NOT mapped: BoardData has no such section (calibration
    // is a parallel path that was never part of an official BoardData load - see
    // KiCadCalibrationDraftWriter). They travel in the manifest for the publisher to handle
    // separately, not through this conversion.
    // ###########################################################################################
    public static class SubmissionBoardData
    {
        public static BoardData FromManifest(SubmissionManifest manifest, string revisionDate = "")
        {
            ArgumentNullException.ThrowIfNull(manifest);

            SubmissionRows rows = manifest.Rows ?? new SubmissionRows();

            return new BoardData
            {
                RevisionDate = revisionDate ?? string.Empty,
                Schematics = [.. rows.Schematics],
                Components = [.. rows.Components],
                ComponentImages = [.. rows.ComponentImages],
                ComponentHighlights = [.. rows.ComponentHighlights],
                ComponentLocalFiles = [.. rows.ComponentLocalFiles],
                ComponentLinks = [.. rows.ComponentLinks],
                BoardLocalFiles = [.. rows.BoardLocalFiles],
                BoardLinks = [.. rows.BoardLinks],
                Credits = [.. rows.Credits],
                KiCadImportantSignals = [.. rows.KiCadImportantSignals]
            };
        }
    }

    // ###########################################################################################
    // What the confirmation dialog shows. AlreadyOnServer is what makes a typo fix visibly cheap:
    // "2 new, 238 already on the server".
    // ###########################################################################################
    public sealed record SubmissionPreview(
        int SchematicCount,
        int ComponentCount,
        int FileCount,
        int NewFileCount,
        int AlreadyOnServer,
        long BytesToUpload);
}
