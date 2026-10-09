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
    // *** THE GENERATION GUARD IS THE POINT OF THIS CLASS. *** The project owner's rule is that
    // publishing writes ONLY the newest workbook generation and NEVER an older one, because older
    // generations are frozen compatibility targets still serving older application builds
    // ("Classic-Repair-Toolbox.xlsx" serves everything before 2.0.0). Writing one is silent
    // damage: the write succeeds, and app builds that depended on that file get contributed data
    // they cannot read, or a file that no longer matches their master workbook. So the target
    // generation is resolved from the tree and every planned workbook path is checked against it.
    //
    // WHAT THIS DOES NOT DO: validate the submission's CONTENT. SubmissionValidator already owns
    // that and runs earlier, twice. This is the last gate, and it asks a different question -
    // not "is this data any good" but "is it safe to write these bytes to these paths".
    // ###########################################################################################
    public static class PublishPlan
    {
        // ###########################################################################################
        // Builds the plan, or returns the reasons it cannot be built.
        //
        // dataRoot         - the tree being published INTO (BETA). Every written path is contained.
        // boardFolder     - the board's folder inside that tree, already resolved by the caller.
        // existingFileNames- the file names ALREADY in the board's folder, used to resolve which
        //                    generation the tree is on. Names only; this class reads no disk.
        // boardStem        - the board workbook's version-free stem ("Data C64 250407"). For a new
        //                    board the caller derives it; for an existing one it comes off the
        //                    file already there, because board file names do not follow the folder
        //                    names mechanically ("Data C128DCR 250477" lives under C128/250477).
        // tree             - what is published NOW (security review, 2026-09-25). Lets a file
        //                    that belongs to another board be cited unchanged, and catches a path
        //                    that differs from a published one only by capitalisation. Null means
        //                    it could not be consulted: every foreign file is then refused, and
        //                    case variants go unchecked. The server always passes the real tree.
        // ###########################################################################################
        public static PublishPlanResult Build(
            string dataRoot,
            string boardFolder,
            IEnumerable<string> existingFileNames,
            string boardStem,
            SubmissionManifest manifest,
            string revision,
            DateTimeOffset publishedUtc,
            IEnumerable<string>? maintainers,
            string origin,
            PublishedTreeView? tree = null)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            var problems = new List<ValidationFinding>();

            if (string.IsNullOrWhiteSpace(dataRoot))
                problems.Add(PublishPlan.Error("publish.no-data-root", string.Empty, "The data root to publish into is not set."));

            if (string.IsNullOrWhiteSpace(boardFolder))
                problems.Add(PublishPlan.Error("publish.no-board-folder", string.Empty, "The board folder to publish into is not set."));

            if (string.IsNullOrWhiteSpace(boardStem))
                problems.Add(PublishPlan.Error("publish.no-workbook-name", string.Empty, "The board workbook name is not known, so there is nothing to write."));

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
            }

            if (problems.Count > 0)
                return PublishPlanResult.Refused(problems);

            // ---- The generation the tree is on -------------------------------------------------
            //
            // DISCOVERED, never configured - see DataGenerationRules. A configured value is a
            // second place the truth lives, and when it falls behind the tree the publish lands in
            // a frozen generation with nothing failing.
            // The names as a list, because two different questions are asked of them below and
            // the caller may hand over a lazy enumerable.
            string[] existing = (existingFileNames ?? []).ToArray();

            Version? targetGeneration = DataGenerationRules.ResolveNewestGeneration(existing);

            // *** "NO GENERATION" HAS TWO MEANINGS AND THEY ARE NOT THE SAME, which the first
            // version of this guard got wrong and an existing test caught. ***
            //
            //   - The folder holds an UNVERSIONED board and nothing newer. That board is ON the
            //     original generation, legitimately, and must keep publishing into it. Refusing
            //     would make every such board unpublishable.
            //
            //   - The folder is EMPTY. There is no generation to read because the board does not
            //     exist yet, and falling back to "no version suffix" would publish
            //     "Data C64 250407.xlsx" - the frozen file serving every pre-2.0.0 build, the one
            //     file the project owner's rule says is NEVER written.
            //
            // ResolveNewestGeneration answers null for both, so the folder's EMPTINESS is what
            // distinguishes them. Only a genuinely new board falls through to the tree.
            bool isNewBoard = existing.Length == 0;

            if (targetGeneration is null && isNewBoard)
            {
                targetGeneration = DataGenerationRules.ResolveNewestGenerationFromTree(dataRoot);

                if (targetGeneration is null)
                {
                    // No versioned master anywhere in the tree. Publishing would have to write the
                    // unversioned file, so it is refused outright rather than guessed at.
                    return PublishPlanResult.Refused(
                    [
                        PublishPlan.Error(
                            "publish.no-generation",
                            boardStem,
                            "This board has no published files and the data tree carries no versioned master workbook, " +
                            "so there is no generation to publish into. Publishing would write the unversioned tree, which is frozen.")
                    ]);
                }
            }

            string workbookFileName = DataGenerationRules.BuildBoardFileName(boardStem, targetGeneration);

            if (string.IsNullOrWhiteSpace(workbookFileName))
            {
                return PublishPlanResult.Refused(
                    [PublishPlan.Error("publish.no-workbook-name", boardStem, $"A workbook name could not be built from [{boardStem}].")]);
            }

            if (!SubmissionPathRules.TryResolve(
                    boardFolder,
                    workbookFileName,
                    out string workbookPath,
                    out string workbookFailure))
            {
                return PublishPlanResult.Refused(
                    [PublishPlan.Error("publish.workbook-path", workbookFileName, $"The board workbook path is refused: {workbookFailure}")]);
            }

            // ---- The board folder's own spelling ----------------------------------------------
            //
            // A "new board" whose folder differs from a published one only by capitalisation is
            // two folders on this server and ONE on every Windows and macOS client, so its files
            // would replace the real board's on their disks. See PublishedTreeView's header.
            string boardRelative = $"{manifest.Manufacturer}/{manifest.Hardware}/{manifest.Board}";
            string? boardVariant = tree?.FindCaseVariant(boardRelative);

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
            var seenPaths = new HashSet<string>(StringComparer.Ordinal);
            var seenPathsIgnoringCase = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // The files the rows actually name - the only ones this publish may carry. Computed
            // once, through the same collector the client and the create-time check use.
            IReadOnlySet<string> referenced = SubmissionFileRules.ReferencedFiles(manifest);

            foreach (SubmissionFile file in manifest.Files ?? [])
            {
                // ###########################################################################################
                // *** RESOLVED AGAINST THE DATA ROOT, NOT THE BOARD FOLDER (fixed 2026-09-23). ***
                //
                // A submitted file path is ALREADY data-root-relative - "Commodore/C64/250407/Board
                // Layout 250407 NTSC.png" - because that is how a board stores its references and
                // how the desktop app resolves them. Resolving against the board folder therefore
                // wrote every file to
                // "<root>/Commodore/C64/250407/Commodore/C64/250407/Board Layout...png": the entire
                // board duplicated INSIDE ITSELF.
                //
                // Reported by the project owner after the first real publish. 1,215 files were written
                // into "250407/Commodore/" and "250407/Generic shared files/", the manifest grew by
                // ~1,200 entries, and every client then downloaded the whole board again - which is
                // how it was noticed at all. The originals were also updated, so the board still
                // WORKED; the duplicates were pure junk that syncs to every user.
                //
                // It could never have been right for a SHARED file either: "Commodore/Shared
                // files/Component images/6526.png" belongs beside the manufacturer, not inside one
                // board - the same reasoning that fixed ReviewAssetLocator's published-file lookup
                // earlier today.
                //
                // CONTAINMENT IS UNCHANGED IN STRENGTH, only rebased: the resolve still refuses
                // anything escaping the root it is given. What it no longer does is confine a
                // publish to one board's folder - which was never the real guard anyway, since
                // SubmissionValidator already refuses a submission whose paths do not belong to it.
                // ###########################################################################################
                if (!SubmissionPathRules.TryResolve(
                        dataRoot,
                        file.Path,
                        out string resolved,
                        out string failure))
                {
                    problems.Add(PublishPlan.Error("file.path-refused", file.Path, $"[{file.Path}] cannot be published: {failure}"));
                    continue;
                }

                if (!SubmissionPathRules.IsValidHash(file.Sha256))
                {
                    problems.Add(PublishPlan.Error("file.hash-invalid", file.Path, $"[{file.Path}] does not carry a usable content hash."));
                    continue;
                }

                // Case-SENSITIVE duplicate: the same path twice is a broken manifest.
                if (!seenPaths.Add(file.Path))
                {
                    problems.Add(PublishPlan.Error("file.duplicate", file.Path, $"[{file.Path}] appears more than once in the submission."));
                    continue;
                }

                // Case-INSENSITIVE collision: both can exist on the Linux server, only one on a
                // Windows client, so the published tree would be un-syncable for much of the
                // audience. Named with BOTH spellings, because "already exists" about a file you
                // cannot see is baffling.
                if (!seenPathsIgnoringCase.Add(file.Path))
                {
                    string other = seenPaths.First(path =>
                        string.Equals(path, file.Path, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(path, file.Path, StringComparison.Ordinal));

                    problems.Add(PublishPlan.Error(
                        "file.case-collision",
                        file.Path,
                        $"[{file.Path}] and [{other}] differ only in capitalisation, which cannot both exist on a Windows machine."));
                    continue;
                }

                // *** THE GENERATION GUARD. *** A submission may not carry a workbook of its own
                // at all - the board's rows are the payload and the workbook is generated from
                // them. Anything that LOOKS like a board workbook of a different generation is
                // refused rather than written.
                if (PublishPlan.LooksLikeWorkbook(file.Path))
                {
                    problems.Add(PublishPlan.Error(
                        "file.workbook-upload",
                        file.Path,
                        $"[{file.Path}] is a workbook. The board workbook is generated from the submitted rows and may not be uploaded."));
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
                // The board's own KiCad data (2026-09-26): its own types, cited by no row - the
                // same exemption ValidateManifestFiles grants, from the same rule.
                bool kiCadProjectFile = SubmissionKiCadFiles.IsSubmittable(manifest, file.Path);

                if (!SubmissionFileRules.TryCheckName(file.Path, out _, out string nameReason, kiCadProjectFile))
                {
                    problems.Add(PublishPlan.Error("file.type-refused", file.Path, nameReason));
                    continue;
                }

                if (!kiCadProjectFile && !referenced.Contains(file.Path))
                {
                    problems.Add(PublishPlan.Error(
                        "file.not-used",
                        file.Path,
                        $"[{file.Path}] is not used by any row of the board, so it cannot be published."));
                    continue;
                }

                var planned = new PlannedFile(file.Path, resolved, file.Sha256, file.SizeBytes);

                // Another board's file may be CITED, never CHANGED - see SubmissionFileScope. One
                // that is byte-identical to what is published is not written at all: it is already
                // there, and writing it would only be a chance to get it wrong.
                if (SubmissionFileScopes.Classify(manifest, file.Path) == SubmissionFileScope.Foreign)
                {
                    string? published = tree?.HashOf(file.Path);

                    if (published is null || !string.Equals(published, file.Sha256, StringComparison.Ordinal))
                    {
                        problems.Add(PublishPlan.Error(
                            "file.other-board",
                            file.Path,
                            $"[{file.Path}] belongs to another board, and this submission would change it."));
                        continue;
                    }

                    unchanged.Add(planned);
                    continue;
                }

                string? variant = tree?.FindCaseVariant(file.Path);

                if (variant is not null)
                {
                    problems.Add(PublishPlan.Error(
                        "path.case-collision",
                        file.Path,
                        $"[{file.Path}] differs only in capitalisation from the published [{variant}], " +
                        "which is the same file on Windows and macOS."));
                    continue;
                }

                // ###########################################################################################
                // *** THE BOARD'S OWN AND SHARED FILES ARE LEFT ALONE WHEN THEY ARE ALREADY THERE, BYTE
                // FOR BYTE (code review, 2026-09-25). *** A manifest lists every file the board cites -
                // about 1,200 for the C64 250407 board - so a one-cell typo fix re-verified and
                // rewrote every one of them, touching each file's modified time, which made
                // PublishedFileHashes, WorkbookReadCache and the checksum manifest re-hash the lot.
                // The rule another board's file already followed now covers all of them: identical
                // bytes at the same path are not written, and stay in the content hash.
                //
                // Safe to decide here because the plan is built under the PublishLock the write
                // happens under, so nothing can change the file between this look and the write.
                // ###########################################################################################
                string? publishedHash = tree?.HashOf(file.Path);

                if (publishedHash is not null && string.Equals(publishedHash, file.Sha256, StringComparison.Ordinal))
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
            // The content hash covers the files AND the workbook, because the workbook is part of
            // what a client syncs - a rows-only change produces an identical file list and would
            // otherwise leave the hash unmoved, so nothing would re-download.
            //
            // The workbook's own hash is not known until it has been written (it is generated, not
            // uploaded), so the caller folds it in via PublishPlanDetail.DescriptorWithWorkbook.
            BoardDescriptor descriptor = BoardDescriptorRules.Build(
                manifest.Manufacturer,
                manifest.Hardware,
                manifest.Board,
                revision,
                publishedUtc,
                maintainers,
                origin,
                [.. files.Concat(unchanged).Select(file => new BoardContentEntry(file.RelativePath, file.Sha256))]);

            return PublishPlanResult.Planned(
                new PublishPlanDetail(
                    manifest.BoardId,
                    boardFolder,
                    targetGeneration,
                    workbookFileName,
                    workbookPath,
                    files,
                    descriptor,
                    dataRoot,
                    unchanged));
        }

        // ###########################################################################################
        // Whether a submitted path looks like a board workbook. Deliberately BROAD - any .xlsx
        // anywhere in the submission - rather than trying to recognise the exact naming of a
        // generation.
        //
        // The narrow version was considered and rejected: it would have to decide whether
        // "Data C64 250407 v1.0.0.xlsx" is a board workbook for an older generation (refuse) or an
        // unrelated spreadsheet a contributor attached (allow), and getting that wrong in the
        // permissive direction overwrites a frozen compatibility target. No shipped board
        // references a spreadsheet as an attachment, so the broad rule costs nothing real.
        // ###########################################################################################
        private static bool LooksLikeWorkbook(string path) =>
            path.EndsWith(DataGenerationRules.WorkbookExtension, StringComparison.OrdinalIgnoreCase);

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
    // One file a publish will write: where it came from (the blob hash) and where it lands.
    //
    // RelativePath is kept alongside the absolute one because the descriptor's content hash is
    // computed over relative paths - an absolute path would fold the server's own directory layout
    // into a hash every client compares against.
    // ###########################################################################################
    public sealed record PlannedFile(
        string RelativePath,
        string AbsolutePath,
        string Sha256,
        long SizeBytes);

    public sealed record PublishPlanDetail(
        string BoardId,
        string BoardFolder,
        Version? TargetGeneration,
        string WorkbookFileName,
        string WorkbookPath,
        IReadOnlyList<PlannedFile> Files,
        BoardDescriptor Descriptor,

        // The tree being published into. Carried so the writer can check the path from here down
        // to each file for a symbolic link before writing through it.
        string DataRoot,

        // Files the submission carries byte-identical to what is already published at the same
        // path: another board's file it cites (see SubmissionFileScope), or one of its own or a
        // shared file it did not change. Never written - they are already there - but part of the
        // board's content, so they stay in its content hash.
        IReadOnlyList<PlannedFile> UnchangedFiles)
    {
        // The total bytes this publish will write, excluding the generated workbook. Reported
        // rather than enforced - a size limit belongs at submission time, where the contributor
        // can still do something about it.
        public long TotalFileBytes => this.Files.Sum(file => file.SizeBytes);

        // ###########################################################################################
        // The `.json` SIDECAR beside the workbook - the board's other half, holding every component
        // highlight and every KiCad calibration.
        //
        // DERIVED from the workbook rather than carried as its own field, and derived through the
        // SAME helper the reader uses. Two independently-stored paths are two things that can
        // disagree, and the failure would be a sidecar written where nothing ever looks for it -
        // a board that publishes with all its highlights silently missing.
        // ###########################################################################################
        public string SidecarPath => BoardComponentHighlightStorage.GetJsonPath(this.WorkbookPath);

        public string SidecarFileName => System.IO.Path.GetFileName(this.SidecarPath);

        // ###########################################################################################
        // The descriptor with the GENERATED WORKBOOK folded into its content hash.
        //
        // *** THE EXECUTOR MUST CALL THIS, AND THE FAILURE IF IT DOES NOT IS SILENT. *** The
        // descriptor built by PublishPlan covers the UPLOADED files only, because the workbook is
        // generated from the submitted rows and its hash cannot be known until it has been
        // written. A rows-only change - the commonest contribution there is - uploads no files at
        // all, so its uploaded-file list is byte-identical to the previous publish's. Without the
        // workbook folded in, the content hash would not move, and no client would ever
        // re-download the board that just changed.
        //
        // This exists as a method rather than as an instruction in a comment precisely because a
        // step the caller must remember is a step the caller eventually forgets. It is covered by
        // its own tests.
        //
        // The workbook is entered under its own file name, so it sits in the same ordered,
        // delimited list as every other file and needs no special case in ComputeContentHash.
        //
        // *** THE SIDECAR IS FOLDED IN FOR THE SAME REASON AND IS EQUALLY REQUIRED. *** A board is
        // TWO files: the workbook and the `.json` beside it holding every component highlight and
        // every KiCad calibration. Both are GENERATED at publish time, so neither hash can be
        // known when the plan is built. A submission that only MOVES A HIGHLIGHT changes no row in
        // the workbook and uploads no file - so without the sidecar here its content hash would
        // not move, and no client would ever re-download the board that just changed. That is the
        // same silent failure the workbook argument exists to prevent, one file over.
        // ###########################################################################################
        public BoardDescriptor DescriptorWithWorkbook(string workbookSha256, string sidecarSha256)
        {
            if (!SubmissionPathRules.IsValidHash(workbookSha256))
            {
                throw new ArgumentException(
                    "The generated workbook's content hash is required before a descriptor can be written.",
                    nameof(workbookSha256));
            }

            if (!SubmissionPathRules.IsValidHash(sidecarSha256))
            {
                throw new ArgumentException(
                    "The generated sidecar's content hash is required before a descriptor can be written.",
                    nameof(sidecarSha256));
            }

            List<BoardContentEntry> entries =
            [
                .. this.Files.Concat(this.UnchangedFiles).Select(file => new BoardContentEntry(file.RelativePath, file.Sha256)),
                new BoardContentEntry(this.WorkbookFileName, workbookSha256),
                new BoardContentEntry(this.SidecarFileName, sidecarSha256)
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
