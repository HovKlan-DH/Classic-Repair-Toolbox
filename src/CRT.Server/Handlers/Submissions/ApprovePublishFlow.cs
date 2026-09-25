using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // APPROVING a submission: the whole chain, from a reviewer's decision to a published board
    // (NewContributeStrategy.md Phase 5, tasks 5 and 6).
    //
    // *** THIS IS THE ONLY IRREVERSIBLE OPERATION IN THE SYSTEM. *** Task 7 was struck by the
    // maintainer, so no publish history is retained: this overwrites the published board in place
    // and the only way back is to publish a correction. Every decision below is made on that
    // basis.
    //
    // *** THE ORDER OF THE CHECKS IS THE DESIGN, not incidental. *** Cheap refusals first, so the
    // dangerous work only ever runs after everything that could refuse has:
    //
    //   1. AUTHORITY   - may this account publish at all? (administrator only)
    //   2. EXISTENCE   - is there such a submission?
    //   3. STATE       - is it still undecided? (the double-publish interlock)
    //   4. PAYLOAD     - can its rows actually be read?
    //   5. PLAN        - are all the paths safe, and which generation is being written?
    //   6. WRITE       - only now does anything touch the tree.
    //
    // Every refusal before step 6 leaves the tree and the submission exactly as they were, so a
    // refused approval can simply be retried once whatever was wrong is fixed.
    //
    // *** A MISSING PAYLOAD IS REFUSED, and that check is load-bearing rather than defensive. ***
    // Without the payload there are no rows, and a board built from no rows is EMPTY. Publishing
    // it would delete the system's entire contents - every component, every highlight - over a
    // board that was working, irreversibly. It is the single most damaging thing this class could
    // be made to do, and it would look like an ordinary successful publish.
    //
    // A COORDINATOR, not a worker: PublishMerge decides what the board becomes, PublishPlan
    // decides where it may be written, PublishExecutor writes it, ReviewDecisionRules decides who
    // may ask. This class only sequences them, which is why it has no logic worth unit testing in
    // isolation - the tests drive the real chain against a real temp tree.
    // ###########################################################################################
    public sealed class ApprovePublishFlow
    {
        private readonly PublishExecutor thisExecutor;
        private readonly PublishedBoardReader thisPublishedBoards;
        private readonly ISubmissionStore thisStore;
        private readonly ILogger<ApprovePublishFlow> thisLogger;

        public ApprovePublishFlow(
            PublishExecutor executor,
            PublishedBoardReader publishedBoards,
            ISubmissionStore store,
            ILogger<ApprovePublishFlow> logger)
        {
            this.thisExecutor = executor;
            this.thisPublishedBoards = publishedBoards;
            this.thisStore = store;
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // Approves a submission and publishes it.
        //
        // Returns an outcome rather than throwing, because every caller is an HTTP endpoint that
        // has to turn a failure into a status code and a sentence a reviewer can act on.
        // ###########################################################################################
        public async Task<ApproveOutcome> ApproveAsync(
            long submissionId,
            AccountRecord? account,
            string? dataTreeRoot,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            // ---- 1. Authority ------------------------------------------------------------------
            //
            // Checked FIRST and against the state we have not read yet only in the sense that
            // ReviewDecisionRules takes both - the role half refuses before anything is loaded.
            if (!ReviewAuthority.CanPublish(account))
            {
                return ApproveOutcome.Refused(
                    "This account is not allowed to publish. Ask an administrator to approve it.",
                    isForbidden: true);
            }

            if (string.IsNullOrWhiteSpace(dataTreeRoot))
            {
                // A missing setting must never make the service's own working directory the data
                // tree - that is where its configuration and binaries live.
                this.thisLogger.LogError(
                    "Approval of submission {SubmissionId} refused: no data tree root is configured.",
                    submissionId);

                return ApproveOutcome.Refused("The server has no data tree configured to publish into.");
            }

            // ---- 2. Existence ------------------------------------------------------------------
            SubmissionRecord? record = await this.thisStore
                .FindAsync(submissionId, cancellationToken)
                .ConfigureAwait(false);

            if (record is null)
                return ApproveOutcome.NotFound();

            // ---- 3. State - the double-publish interlock ---------------------------------------
            //
            // Re-read from the store rather than trusted from the caller: the app's belief that a
            // submission is still pending can be seconds out of date, and with no revision history
            // a second publish is not a harmless no-op.
            if (!ReviewDecisionRules.CanApprove(account, record.State, out string why))
                return ApproveOutcome.Refused(why, isConflict: true);

            // ---- 4. The payload ----------------------------------------------------------------
            //
            // See the class header: an unreadable payload means no rows, and publishing no rows
            // deletes the board.
            SubmissionManifest? manifest = await this.thisStore
                .LoadPayloadAsync(submissionId, cancellationToken)
                .ConfigureAwait(false);

            if (manifest is null)
            {
                this.thisLogger.LogError(
                    "Approval of submission {SubmissionId} refused: its payload could not be loaded, " +
                    "so publishing would write an empty board over a working one.",
                    submissionId);

                return ApproveOutcome.Refused(
                    "This submission's contents could not be loaded, so it cannot be published.");
            }

            // ---- 5. The plan -------------------------------------------------------------------
            //
            // Decides every path and refuses every unsafe one BEFORE a byte is written, which is
            // what makes an irreversible operation safe to start.
            BoardData? published = await this.thisPublishedBoards
                .TryReadAsync(dataTreeRoot, manifest, cancellationToken)
                .ConfigureAwait(false);

            BoardData board = PublishMerge.Build(manifest, published);

            PublishPlanResult plan = ApprovePublishFlow.BuildPlan(
                dataTreeRoot, manifest, board, published, nowUtc);

            if (!plan.IsPlanned)
            {
                string reasons = string.Join(" ", plan.Problems.Select(problem => problem.Message));

                this.thisLogger.LogWarning(
                    "Approval of submission {SubmissionId} refused by the publish plan: {Reasons}",
                    submissionId, reasons);

                return ApproveOutcome.Refused($"This submission cannot be published: {reasons}");
            }

            // ---- 6. Write --------------------------------------------------------------------
            //
            // The first step that touches the tree. Everything above can refuse for free.
            PublishOutcome outcome = await this.thisExecutor
                .ExecuteAsync(
                    plan.Plan!,
                    board,
                    PublishMerge.CalibrationsOf(manifest),
                    submissionId,
                    nowUtc,
                    cancellationToken)
                .ConfigureAwait(false);

            if (!outcome.IsPublished)
            {
                // The submission stays PENDING deliberately, so it can be retried once whatever
                // stopped it is fixed - a missing blob being the likely cause. Marking it failed
                // would need somebody to resurrect it by hand.
                return ApproveOutcome.Refused(outcome.Failure ?? "The publish did not complete.");
            }

            // The decision row, recording WHO approved it. PublishExecutor already moved the state
            // to merged; this adds the reviewer and the audit trail, which SetStateAsync cannot.
            await this.thisStore
                .SetDecisionAsync(
                    submissionId,
                    SubmissionState.Merged,
                    account!.Id,
                    comment: null,
                    nowUtc,
                    cancellationToken)
                .ConfigureAwait(false);

            this.thisLogger.LogInformation(
                "Submission {SubmissionId} was published to {SystemId} at revision {Revision} by {Account}.",
                submissionId, manifest.SystemId, outcome.Descriptor!.Revision, account.Email);

            return ApproveOutcome.Published(outcome.Descriptor);
        }

        // ###########################################################################################
        // The publish plan for this submission.
        //
        // *** ORIGIN IS SET ONCE AND NEVER RECOMPUTED. *** A system that arrived through this
        // pipeline stays "contributed" however many times it is later revised, including by the
        // maintainer - SystemDescriptorRules says so outright. So an EXISTING descriptor's origin
        // wins, and only a system with none at all is classified here.
        //
        // The revision is the board's own REVISION DATE - see PublishMerge.RevisionOf for why that
        // is a contract with the client rather than a choice.
        // ###########################################################################################
        private static PublishPlanResult BuildPlan(
            string dataTreeRoot,
            SubmissionManifest manifest,
            BoardData board,
            BoardData? published,
            DateTimeOffset nowUtc)
        {
            PublishedBoardLocation location = PublishedBoardLocator.Locate(dataTreeRoot, manifest);

            string systemFolder = location.Exists
                ? location.SystemFolder
                : Path.Combine(
                    dataTreeRoot,
                    manifest.Manufacturer.Trim(),
                    manifest.Hardware.Trim(),
                    manifest.Board.Trim());

            // What is already on disk, so the plan can spot a case-only collision against a file
            // it is not itself writing.
            IEnumerable<string> existing = Directory.Exists(systemFolder)
                ? Directory.EnumerateFiles(systemFolder)
                    .Select(Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Select(name => name!)
                : [];

            // Read rather than assumed, so a system's origin and maintainers survive a republish.
            SystemDescriptor? current = SystemDescriptorStore.Read(systemFolder);

            return PublishPlan.Build(
                dataTreeRoot,
                systemFolder,
                existing,
                ApprovePublishFlow.BoardStem(manifest, location),
                manifest,
                PublishMerge.RevisionOf(board),
                nowUtc,
                current?.Maintainers,
                current?.Origin is { Length: > 0 } origin
                    ? origin
                    : SystemDescriptorRules.SystemOrigin.Contributed);
        }

        // ###########################################################################################
        // The board file's stem - "Data C64 250407", the part before the generation suffix.
        //
        // *** READ OFF THE EXISTING FILE WHERE THERE IS ONE, because it does NOT follow the folder
        // names mechanically. *** "Data C128DCR 250477" lives under C128/250477, so rebuilding it
        // from the identity would write a second, differently-named board beside the real one and
        // the system would then carry two. Only a genuinely new system falls back to the
        // convention.
        // ###########################################################################################
        private static string BoardStem(SubmissionManifest manifest, PublishedBoardLocation location)
        {
            if (location.Exists)
            {
                string name = Path.GetFileNameWithoutExtension(location.WorkbookPath);

                // Strip the generation suffix, leaving the stem the plan will re-add one to.
                Version? generation = DataGenerationRules.TryReadGeneration(
                    Path.GetFileName(location.WorkbookPath));

                if (generation is not null)
                {
                    int suffix = name.LastIndexOf(" v", StringComparison.Ordinal);

                    if (suffix > 0)
                        name = name[..suffix];
                }

                if (!string.IsNullOrWhiteSpace(name))
                    return name;
            }

            return $"Data {manifest.Hardware.Trim()} {manifest.Board.Trim()}";
        }
    }

    // ###########################################################################################
    // What approving produced, or why it did not.
    //
    // IsForbidden and IsConflict are distinguished because the endpoint turns them into different
    // status codes, and the difference matters to the reader: "your account may not do this" and
    // "somebody else already decided this" send a reviewer to completely different places.
    // ###########################################################################################
    public sealed record ApproveOutcome(
        bool IsPublished,
        string Error,
        SystemDescriptor? Descriptor,
        bool IsNotFound = false,
        bool IsForbidden = false,
        bool IsConflict = false)
    {
        public static ApproveOutcome Published(SystemDescriptor descriptor) =>
            new(true, string.Empty, descriptor);

        public static ApproveOutcome Refused(
            string error, bool isForbidden = false, bool isConflict = false) =>
            new(false, error, null, IsForbidden: isForbidden, IsConflict: isConflict);

        public static ApproveOutcome NotFound() =>
            new(false, "No such submission.", null, IsNotFound: true);
    }
}
