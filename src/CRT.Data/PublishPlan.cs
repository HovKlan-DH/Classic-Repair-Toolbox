using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Decides WHAT a publish will do, before anything is written (NewContributeStrategy.md
    // Phase 5, task 6). Pure: a submission manifest and the target tree's facts go in, a plan or a
    // list of refusals comes out. Nothing here opens a file.
    //
    // *** WHY A PLAN OBJECT RATHER THAN JUST DOING IT. *** Publishing is the one operation in this
    // board that changes what every user downloads, and it is NOT reversible - the project owner
    // decided against retained revisions (open question 5), so there is no previous version to go
    // back to. An irreversible operation deserves to be decidable and inspectable in full before
    // its first byte lands: every refusal is found up front rather than halfway through a write
    // that has already replaced half a board.
    //
    // *** WHAT A PUBLISH WRITES IN THE CRT 4.0 TREE (2026-10-25). *** Two kinds of file, both at
    // places the SERVER decides - never a path a submission sent:
    //
    //   - the board's own file, "<board id>/board.json" (DataTreeLayout.BoardFilePath), written
    //     from the rows with the board's list of files by name;
    //   - each file's bytes, stored ONCE under their checksum, "Files/<ab>/<sha256>.<ext>"
    //     (DataTreeLayout.StoredFilePath). A stored file never changes - the same name means the
    //     same bytes - so one the tree already holds is not written at all (`isStored`), and a
    //     changed file is new bytes beside the old.
    //
    // A submitted file's `Path` is the board's NAME for it, the name its rows cite. It is still
    // checked as a path (SubmissionPathRules): CRT puts a file under its name in a draft and a KiCad
    // copy, so a name escaping its folder would escape there.
    //
    // WHAT THIS DOES NOT DO: validate the submission's CONTENT. SubmissionValidator already owns
    // that and runs earlier, twice. This is the last gate, and it asks a different question -
    // not "is this data any good" but "is it safe to write these bytes to these places".
    // ###########################################################################################
    public static class PublishPlan
    {
        // ###########################################################################################
        // Builds the plan, or returns the reasons it cannot be built.
        //
        // dataRoot  - the tree being published INTO (BETA). Every written path is contained.
        // board     - what is published NOW (security review, 2026-09-25): the published board's
        //             files by name, and the tree's folders - which catch a NEW board's folder that
        //             differs from a published one only by capitalisation. Null means it could not be
        //             consulted: the folder goes unchecked, and no straggler of the KiCad folder is
        //             admitted for being published already (SubmissionKiCadFiles).
        // isStored  - whether the tree already holds a stored file (data-root-relative); null: none.
        // ###########################################################################################
        public static PublishPlanResult Build(
            string dataRoot,
            SubmissionManifest manifest,
            string revision,
            DateTimeOffset publishedUtc,
            IEnumerable<string>? maintainers,
            string origin,
            PublishedTreeView? board = null,
            Func<string, bool>? isStored = null)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            var problems = new List<ValidationFinding>();

            if (string.IsNullOrWhiteSpace(dataRoot))
                problems.Add(PublishPlan.Error("publish.no-data-root", string.Empty, "The data root to publish into is not set."));

            if (string.IsNullOrWhiteSpace(revision))
                problems.Add(PublishPlan.Error("publish.no-revision", string.Empty, "A publish must carry a revision."));

            // A wrong or missing board id means everything keyed off it - the database row, the
            // folder, every future submission's base - disagrees with what is being written.
            if (!BoardDescriptorRules.IsValidBoardId(manifest.BoardId))
            {
                problems.Add(PublishPlan.Error(
                    "board.id-invalid",
                    manifest.BoardId,
                    $"The submission's board id [{manifest.BoardId}] is not a valid board id."));
            }
            else
            {
                string expected = BoardDescriptorRules.BuildBoardId(
                    manifest.Manufacturer,
                    manifest.Hardware,
                    manifest.Board);

                if (!string.Equals(manifest.BoardId, expected, StringComparison.Ordinal))
                {
                    problems.Add(PublishPlan.Error(
                        "board.id-mismatch",
                        manifest.BoardId,
                        $"The submission's board id [{manifest.BoardId}] does not match the names it carries [{expected}]."));
                }
                else if (DataTreeLayout.ReservedTopFolders.Contains(manifest.Manufacturer.Trim(), StringComparer.OrdinalIgnoreCase))
                {
                    // The tree's own folders - a board there would sit among the stored files.
                    problems.Add(PublishPlan.Error(
                        "board.reserved-folder",
                        manifest.BoardId,
                        $"[{manifest.Manufacturer}] is a folder of the data itself and cannot be a manufacturer."));
                }
            }

            if (problems.Count > 0)
                return PublishPlanResult.Refused(problems);

            if (!SubmissionPathRules.TryResolve(
                    dataRoot,
                    DataTreeLayout.BoardFilePath(manifest.BoardId),
                    out string boardFilePath,
                    out string boardFileFailure))
            {
                return PublishPlanResult.Refused(
                    [PublishPlan.Error("publish.board-file-path", manifest.BoardId, $"The board file path is refused: {boardFileFailure}")]);
            }

            string boardFolder = System.IO.Path.GetDirectoryName(boardFilePath)!;

            // ---- The board folder's own spelling ----------------------------------------------
            //
            // A "new board" whose folder differs from a published one only by capitalisation is
            // two folders on this server and ONE on every Windows and macOS client, so its board
            // file would replace the real board's on their disks. See PublishedTreeView's header.
            string boardRelative = $"{manifest.Manufacturer}/{manifest.Hardware}/{manifest.Board}";
            string? boardVariant = board?.FindCaseVariant(boardRelative);

            if (boardVariant is not null)
            {
                return PublishPlanResult.Refused(
                [
                    PublishPlan.Error(
                        "board.case-collision",
                        boardRelative,
                        $"[{boardRelative}] differs only in capitalisation from the published [{boardVariant}], " +
                        "which would be the same folder on Windows and macOS.")
                ]);
            }

            // ---- The files ---------------------------------------------------------------------
            var files = new List<PlannedFile>();
            var unchanged = new List<PlannedFile>();
            var seenNames = new HashSet<string>(StringComparer.Ordinal);
            var seenNamesIgnoringCase = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // The files the rows actually name - the only ones besides the KiCad data this publish
            // may carry. Computed once, through the same collector the client and the create-time
            // check use.
            IReadOnlySet<string> referenced = SubmissionFileRules.ReferencedFiles(manifest);

            foreach (SubmissionFile file in manifest.Files ?? [])
            {
                if (!SubmissionPathRules.TryResolve(boardFolder, file.Path, out _, out string failure))
                {
                    problems.Add(PublishPlan.Error("file.path-refused", file.Path, $"[{file.Path}] cannot be published: {failure}"));
                    continue;
                }

                if (!SubmissionPathRules.IsValidHash(file.Sha256))
                {
                    problems.Add(PublishPlan.Error("file.hash-invalid", file.Path, $"[{file.Path}] does not carry a usable content hash."));
                    continue;
                }

                // Case-SENSITIVE duplicate: the same name twice is a broken manifest.
                if (!seenNames.Add(file.Path))
                {
                    problems.Add(PublishPlan.Error("file.duplicate", file.Path, $"[{file.Path}] appears more than once in the submission."));
                    continue;
                }

                // Case-INSENSITIVE collision: two names one board cannot hold apart - CRT puts a file
                // under its name in a draft and a KiCad copy, and on Windows and macOS the two are
                // one. Named with BOTH spellings.
                if (!seenNamesIgnoringCase.Add(file.Path))
                {
                    string other = seenNames.First(name =>
                        string.Equals(name, file.Path, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(name, file.Path, StringComparison.Ordinal));

                    problems.Add(PublishPlan.Error(
                        "file.case-collision",
                        file.Path,
                        $"[{file.Path}] and [{other}] differ only in capitalisation, which cannot both exist on a Windows machine."));
                    continue;
                }

                // ###########################################################################################
                // *** THE SAME FILE RULES AS AT CREATE, AGAIN, HERE (security review, 2026-09-25). ***
                //
                // SubmissionFileRules already refused all of these before anything was uploaded. They
                // are repeated because this is the last gate before an irreversible write, and a
                // submission queued by an older build - or a rule changed since - must not reach the
                // tree on the strength of a check that ran weeks ago.
                // ###########################################################################################
                if (!SubmissionFileRules.TryAdmit(manifest, file, referenced, board, out string code, out string reason))
                {
                    problems.Add(code == "file.not_used"
                        ? PublishPlan.Error("file.not-used", file.Path, $"[{file.Path}] is not used by any row of the board, so it cannot be published.")
                        : PublishPlan.Error("file.type-refused", file.Path, reason));
                    continue;
                }

                string stored = DataTreeLayout.StoredFilePath(file.Sha256, file.Path);

                if (!SubmissionPathRules.TryResolve(dataRoot, stored, out string storedPath, out string storedFailure))
                {
                    problems.Add(PublishPlan.Error("file.path-refused", file.Path, $"[{file.Path}] cannot be stored: {storedFailure}"));
                    continue;
                }

                var planned = new PlannedFile(file.Path, storedPath, file.Sha256, file.SizeBytes);

                // ###########################################################################################
                // *** BYTES THE TREE ALREADY HOLDS ARE NOT WRITTEN AGAIN (code review, 2026-09-25). ***
                // A manifest lists every file of the board - about 1,200 for the C64 250407 board - and a
                // one-cell typo fix must not rewrite them all. A stored file is named by its checksum,
                // so one that is there IS these bytes; it stays in the board's list and content hash.
                //
                // Safe to decide here because the plan is built under the PublishLock the write happens
                // under, so nothing can remove the file between this look and the write.
                // ###########################################################################################
                if (isStored?.Invoke(stored) == true)
                {
                    unchanged.Add(planned);
                    continue;
                }

                files.Add(planned);
            }

            if (problems.Count > 0)
                return PublishPlanResult.Refused(problems);

            // ---- The descriptor ----------------------------------------------------------------
            //
            // The content hash covers the board's files by name AND its board file, because the
            // board file is part of what a client syncs - a rows-only change produces an identical
            // file list and would otherwise leave the hash unmoved, so nothing would re-download.
            // The board file's own hash is not known until it has been written (it is generated,
            // not uploaded), so the executor folds it in via PublishPlanDetail.DescriptorWithBoardFile.
            BoardDescriptor descriptor = BoardDescriptorRules.Build(
                manifest.Manufacturer,
                manifest.Hardware,
                manifest.Board,
                revision,
                publishedUtc,
                maintainers,
                origin,
                [.. files.Concat(unchanged).Select(file => new BoardContentEntry(file.Name, file.Sha256))]);

            return PublishPlanResult.Planned(
                new PublishPlanDetail(
                    manifest.BoardId,
                    boardFolder,
                    boardFilePath,
                    files,
                    descriptor,
                    dataRoot,
                    unchanged));
        }

        // Codes follow SubmissionValidator's own "subject.problem" convention, so a caller can
        // group or translate a publish refusal exactly as it groups a validation finding.
        private static ValidationFinding Error(string code, string subject, string message) => new()
        {
            Severity = ValidationSeverity.Error,
            Code = code,
            Subject = subject,
            Message = message
        };
    }

    // ###########################################################################################
    // One file of the published board: its name, its bytes (the blob hash), and where those bytes
    // are stored in the tree. Several names may share one stored file - the same bytes.
    // ###########################################################################################
    public sealed record PlannedFile(
        string Name,
        string AbsolutePath,
        string Sha256,
        long SizeBytes)
    {
        // "Files/3f/3f7a...e1.png" - data-root-relative, for the manifest and the logs.
        public string StoredPath => DataTreeLayout.StoredFilePath(this.Sha256, this.Name);
    }

    public sealed record PublishPlanDetail(
        string BoardId,
        string BoardFolder,

        // "<data root>/<board id>/board.json" - absolute, already contained.
        string BoardFilePath,

        // The files whose bytes this publish stores - the tree does not hold them yet.
        IReadOnlyList<PlannedFile> Files,
        BoardDescriptor Descriptor,

        // The tree being published into. Carried so the writer can check the path from here down
        // to each file for a symbolic link before writing through it.
        string DataRoot,

        // The files whose bytes the tree already holds. Never written - they are already there -
        // but the board's files all the same, so they stay in its list and content hash.
        IReadOnlyList<PlannedFile> UnchangedFiles)
    {
        // The total bytes this publish will store, excluding the generated board file. Reported
        // rather than enforced - a size limit belongs at submission time, where the contributor
        // can still do something about it.
        public long TotalFileBytes => this.StoredFiles.Sum(file => file.SizeBytes);

        // "<board id>/board.json", data-root-relative.
        public string BoardFileRelative => DataTreeLayout.BoardFilePath(this.BoardId);

        // The files to store, each stored file once (two names of the same bytes are one copy).
        public IReadOnlyList<PlannedFile> StoredFiles =>
            this.Files.DistinctBy(file => file.StoredPath, StringComparer.Ordinal).ToList();

        // ###########################################################################################
        // The board's list of files, as its board file is to hold it: every name with its bytes,
        // in name order - DraftFiles' order, so a draft and the published board list alike.
        // ###########################################################################################
        public List<BoardDocument.StoredFileRow> FileList =>
            this.Files.Concat(this.UnchangedFiles)
                .Select(file => new BoardDocument.StoredFileRow { Name = file.Name, Sha256 = file.Sha256, Size = file.SizeBytes })
                .OrderBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(file => file.Name, StringComparer.Ordinal)
                .ToList();

        // ###########################################################################################
        // The descriptor with the GENERATED BOARD FILE folded into its content hash.
        //
        // *** THE EXECUTOR MUST CALL THIS, AND THE FAILURE IF IT DOES NOT IS SILENT. *** The
        // descriptor built by PublishPlan covers the files only, because the board file is
        // generated from the submitted rows and its hash cannot be known until it has been written.
        // A rows-only change - the commonest contribution there is - stores no file at all, so its
        // file list is byte-identical to the previous publish's. Without the board file folded in,
        // the content hash would not move, and the production promotion - which compares exactly
        // this hash to decide what is waiting - would never offer the board. The board file holds
        // the highlights, calibrations and renames too, so a submission that only MOVES A HIGHLIGHT
        // moves it as well.
        // ###########################################################################################
        public BoardDescriptor DescriptorWithBoardFile(string boardFileSha256)
        {
            if (!SubmissionPathRules.IsValidHash(boardFileSha256))
            {
                throw new ArgumentException(
                    "The generated board file's content hash is required before a descriptor can be written.",
                    nameof(boardFileSha256));
            }

            List<BoardContentEntry> entries =
            [
                .. this.Files.Concat(this.UnchangedFiles).Select(file => new BoardContentEntry(file.Name, file.Sha256)),
                new BoardContentEntry(DataTreeLayout.BoardFileName, boardFileSha256)
            ];

            // A NEW descriptor rather than a mutation of the planned one: BoardDescriptor is a
            // mutable class, and a plan that quietly changed under a caller holding it would be
            // the sort of surprise this whole "decide it all up front" design exists to avoid.
            return new BoardDescriptor
            {
                BoardId = this.Descriptor.BoardId,
                Manufacturer = this.Descriptor.Manufacturer,
                Hardware = this.Descriptor.Hardware,
                Board = this.Descriptor.Board,
                Revision = this.Descriptor.Revision,
                PublishedUtc = this.Descriptor.PublishedUtc,
                Maintainers = [.. this.Descriptor.Maintainers],
                Origin = this.Descriptor.Origin,
                ContentHash = BoardDescriptorRules.ComputeContentHash(this.Descriptor.Revision, entries)
            };
        }
    }

    // ###########################################################################################
    // Either a plan or the reasons there isn't one. Deliberately not "a plan plus a problems
    // list": a partially valid publish is not a thing that may proceed, and a caller holding both
    // would have to remember to check.
    // ###########################################################################################
    public sealed class PublishPlanResult
    {
        private PublishPlanResult(PublishPlanDetail? plan, IReadOnlyList<ValidationFinding> problems)
        {
            this.Plan = plan;
            this.Problems = problems;
        }

        public PublishPlanDetail? Plan { get; }

        public IReadOnlyList<ValidationFinding> Problems { get; }

        public bool IsPlanned => this.Plan != null;

        public static PublishPlanResult Planned(PublishPlanDetail plan) => new(plan, []);

        public static PublishPlanResult Refused(IReadOnlyList<ValidationFinding> problems) =>
            new(null, problems);
    }
}
