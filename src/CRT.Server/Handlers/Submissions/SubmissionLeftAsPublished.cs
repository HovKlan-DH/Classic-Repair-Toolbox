using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A FILE THE CONTRIBUTOR NEVER TOUCHED STAYS AS IT IS PUBLISHED NOW (code review, 2026-10-13).
    //
    // A submission carries every file its board cites, each with the hash the contributor's draft
    // had - the draft is a whole copy of the board. The approval's plan (PublishPlan) writes every
    // file whose published hash differs from the submission's, and for a file the contributor left
    // alone that difference means somebody ELSE changed it after the submission was created. The
    // reviewer's case: S1 cites a shared image unchanged, S2 replaces the image and is published,
    // S1 is approved - and S1's copy of the OLD image went back over S2's, for every board citing it.
    //
    // At create the SERVER records which files were byte-identical to the published tree
    // (NewSubmission.LeftAsPublished, migration 0021). Before anything reads the submission's files
    // for a decision - the approval, the shared-file test, the file tree - such a file whose
    // published copy has moved since is pointed at the bytes published NOW. The plan then finds it
    // equal and does not write it, the shared-file test no longer calls it a replacement, and the
    // file tree shows it unchanged: one rule, applied once, before every reader.
    //
    // A file the contributor DID change is untouched by this - its hash never matched the published
    // one at create, so it has no row. An amended file whose hash differs from its row is the
    // maintainer's change, and is written as one. The stored payload is never rewritten.
    // ###########################################################################################
    public static class SubmissionLeftAsPublished
    {
        // ###########################################################################################
        // The files of `manifest` that were byte-identical to a published tree at create - the BETA
        // tree `tree`, or the STABLE tree `stableTree`: the record CreateAsync stores. A file with
        // nothing published at its path, or other bytes in both, is not one. No tree, no record -
        // the safe direction: such a file is written as submitted, which is what happened before
        // this record existed.
        //
        // *** STABLE TOO (code review, 2026-10-13). *** A contributor's draft is a copy of the STABLE
        // source by default. With S2's newer image already in BETA, a draft from stable cites the
        // OLD image unchanged - which differs from BETA, so judged against BETA alone it was "changed"
        // and the approval wrote it back over S2's. Matching either published copy is "left alone".
        // ###########################################################################################
        public static IReadOnlyList<SubmissionFile> AtCreate(SubmissionManifest manifest, PublishedTreeView? tree, PublishedTreeView? stableTree = null)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            if (tree is null && stableTree is null)
                return [];

            return manifest.Files
                .Where(file => !string.IsNullOrWhiteSpace(file.Path) &&
                               (string.Equals(tree?.HashOf(file.Path), file.Sha256, StringComparison.Ordinal) ||
                                string.Equals(stableTree?.HashOf(file.Path), file.Sha256, StringComparison.Ordinal)))
                .DistinctBy(file => file.Path, StringComparer.Ordinal)
                .ToList();
        }

        // ###########################################################################################
        // `manifest` with every file left as published at create - `leftAsPublished`, path to the
        // hash it had then - whose published copy has moved since, carrying the published copy's
        // hash AND size instead (the size code review, 2026-10-13: it kept the submitted size beside
        // the new hash, and the file tree showed the old file's size). The same object when nothing
        // moved. `kept` names the files so treated.
        // ###########################################################################################
        public static SubmissionManifest KeepNewerPublished(
            SubmissionManifest manifest,
            IReadOnlyDictionary<string, string>? leftAsPublished,
            PublishedTreeView? tree,
            out IReadOnlyList<string> kept)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            kept = [];

            if (leftAsPublished is null || leftAsPublished.Count == 0 || tree is null)
                return manifest;

            var files = new List<SubmissionFile>(manifest.Files.Count);
            var keptPaths = new List<string>();

            foreach (SubmissionFile file in manifest.Files)
            {
                if (leftAsPublished.TryGetValue(file.Path, out string? atCreate) &&
                    string.Equals(atCreate, file.Sha256, StringComparison.Ordinal) &&
                    tree.HashOf(file.Path) is string now &&
                    !string.Equals(now, file.Sha256, StringComparison.Ordinal))
                {
                    files.Add(new SubmissionFile { Path = file.Path, Sha256 = now, SizeBytes = tree.SizeOf(file.Path) ?? file.SizeBytes });
                    keptPaths.Add(file.Path);
                    continue;
                }

                files.Add(file);
            }

            if (keptPaths.Count == 0)
                return manifest;

            kept = keptPaths;

            return new SubmissionManifest
            {
                FormatVersion = manifest.FormatVersion,
                BoardId = manifest.BoardId,
                Manufacturer = manifest.Manufacturer,
                Hardware = manifest.Hardware,
                Board = manifest.Board,
                BaseRevision = manifest.BaseRevision,
                Summary = manifest.Summary,
                HardwareNotes = manifest.HardwareNotes,
                ContactEmail = manifest.ContactEmail,
                ApplicationVersion = manifest.ApplicationVersion,
                CreatedUtc = manifest.CreatedUtc,
                Files = files,
                Rows = manifest.Rows,
                Renames = manifest.Renames
            };
        }

        // ###########################################################################################
        // A submission's payload as every decision about it must read it: loaded, then the rule
        // above applied against the tree at `dataTreeRoot` as it is now. Null when the payload
        // cannot be loaded.
        // ###########################################################################################
        public static async Task<SubmissionManifest?> LoadAsync(
            ISubmissionStore store,
            long submissionId,
            string? dataTreeRoot,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);

            SubmissionManifest? manifest = await store.LoadPayloadAsync(submissionId, cancellationToken).ConfigureAwait(false);

            if (manifest is null)
                return null;

            IReadOnlyDictionary<string, string> left = await store.GetLeftAsPublishedAsync(submissionId, cancellationToken).ConfigureAwait(false);

            return SubmissionLeftAsPublished.KeepNewerPublished(manifest, left, PublishedTreeProbe.For(dataTreeRoot), out _);
        }
    }
}
