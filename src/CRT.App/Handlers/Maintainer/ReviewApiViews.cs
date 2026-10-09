using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // The records ReviewApiParser reads the server's answers into - moved out of ReviewApiParser.cs
    // when it passed the project's ~1,500 lines (code review, 2026-09-27). View types for this
    // application's screens; the records both ends share are CRT.Data's ReviewApiContract.
    // ###########################################################################################

    // IsAdministrator is what shows the "Maintainers" button; the server refuses the screen's
    // requests from anyone else regardless. Trailing with a default so an older server that does
    // not send it reads as "not an administrator", which hides a button rather than a queue.
    public sealed record ReviewQueueResponse(
        bool CanPublish,
        IReadOnlyList<ReviewQueueRow> Submissions,
        bool IsAdministrator = false);

    // ###########################################################################################
    // The administrator's lists (Phase 6 roles). View types, like everything else in this file.
    // ###########################################################################################
    public sealed record ReviewBoardsResponse(IReadOnlyList<ReviewBoardRow> Boards);

    public sealed record ReviewBoardRow(
        string BoardId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? CurrentRevision,
        bool IsAccepting,
        IReadOnlyList<MaintainerRow> Maintainers);

    public sealed record MaintainerRow(long AccountId, string DisplayName, string Email);

    public sealed record ReviewAccountsResponse(IReadOnlyList<ReviewAccountRow> Accounts);

    // ###########################################################################################
    // BETA to production (2026-09-25). View types, apart from the files, which are CRT.Data's own.
    // ###########################################################################################
    public sealed record ProductionListResponse(bool Configured, IReadOnlyList<ProductionBoardRow> Boards);

    public sealed record ProductionBoardRow(
        string BoardId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? BetaRevision,
        string BetaContentHash,
        string? ProductionRevision,
        DateTimeOffset? ProductionPublishedUtc,

        // Whether it waits for THIS account rather than the other approver (2026-09-27) - the BETA
        // badge counts these. Null from an older server, which the badge counts as yours.
        bool? AwaitsYou = null,

        // Whether a contributor whose work it carries discarded their own draft (2026-09-28).
        bool? CarriesDiscardedDraft = null,

        // Whether it waits for the administrator because only administrators publish to stable
        // (2026-10-05) - the row then says so instead of "with the other approver". Null from an
        // older server.
        bool? WaitsForAdministrator = null);

    public sealed record ProductionPlanView(
        string BoardId,
        string? BetaRevision,

        // Sent back with the publish request: the server refuses if BETA moved since.
        string BetaContentHash,
        bool TouchesSharedFiles,
        bool CanPublish,
        string? Refusal,
        int UnchangedCount,
        IReadOnlyList<PromotionFile> Files,
        IReadOnlyList<ReviewFindingView> Problems,
        ApprovalStatus? Approval = null,

        // What promoting would remove from production - shown before anyone approves, and sent
        // back with the publish request.
        FileRemovalPreview? Removals = null,

        // Whose merged work this promotion carries out to everyone (2026-09-27). Empty from an
        // older server.
        IReadOnlyList<CarriedSubmission>? Carrying = null,

        // The files already the same in production - with Files and Removals, the whole file tree
        // (2026-09-28) - and where BETA and production are published, to open a file from it. Empty
        // or null from an older server, which then draws only what changes.
        IReadOnlyList<string>? UnchangedFiles = null,
        string? BetaDataUrl = null,
        string? ProductionDataUrl = null,

        // Every file's size, by path (2026-10-04) - null from an older server, whose tree then shows
        // no sizes.
        IReadOnlyDictionary<string, long>? FileSizes = null);

    // State "awaiting" is a recorded approval that published nothing - the first of two.
    // ###########################################################################################
    // Rolling a BETA board back (2026-09-27). `Returning` is who goes back to the queue - named,
    // because a rollback reverts the whole board and every one of them loses its place in BETA.
    // ###########################################################################################
    //
    // SharedRestored (code review, 2026-09-27): shared files the returning submissions changed, put
    // back to production's bytes. Every board citing a shared file sees it, so the confirmation
    // names them. Empty from a server that does not send them. A rollback never removes a shared
    // file, so there is nothing else to name.
    public sealed record BetaRollbackPlanView(
        string BoardId,
        BetaRollbackKind Kind,
        IReadOnlyList<string> Restored,
        IReadOnlyList<string> Removed,
        IReadOnlyList<CarriedSubmission> Returning,
        IReadOnlyList<string>? SharedRestored = null);

    // Rejected (2026-09-28): the submissions were rejected rather than returned to the queue - false
    // from a server older than Beta > Prod's "Reject", which pushes back instead.
    public sealed record BetaRollbackResult(
        string BoardId,
        BetaRollbackKind Kind,
        int FilesRestored,
        int FilesRemoved,
        int SubmissionsReturned,
        bool Rejected = false);

    public sealed record ProductionPublishResult(
        string BoardId,
        string? Revision,
        int FilesCopied,
        string State = "published",
        IReadOnlyList<ApproverRole>? WaitingFor = null,
        IReadOnlyList<string>? RemovedFiles = null)
    {
        public bool IsAwaitingApproval => string.Equals(this.State, "awaiting", StringComparison.Ordinal);
    }

    public sealed record ReviewAccountRow(
        long Id,
        string Email,
        string DisplayName,
        bool IsAdministrator,
        bool IsVerified,
        bool IsLocked);

    // ###########################################################################################
    // One submission as the review window holds it.
    //
    // *** THESE ARE VIEW TYPES, NOT THE SERVER'S OWN RECORDS. *** CRT.Data's ReviewChangeSummary
    // could have been deserialised directly and that was considered; it was rejected because the
    // server's record carries computed properties and a shape that exists to be COMPUTED, whereas
    // what the window needs is a shape that exists to be DRAWN - sections already filtered to the
    // ones that changed, renames already paired. Sharing the type would also mean any future
    // change to the comparison's internals became a wire-format change by accident.
    //
    // Changes is NULL when the server could not build a summary (an unloadable payload). That is
    // distinct from a summary with no changes, and the window says something different for each.
    // ###########################################################################################
    public sealed record ReviewSubmissionDetail(
        ReviewQueueRow Submission,
        bool CanPublish,
        ReviewChangeSummaryView? Changes,
        IReadOnlyList<ReviewFindingView> Findings,

        // The files this submission carries, for task 4's visual comparison. Never null - an empty
        // list is a rows-only change, which is the commonest contribution there is, and a null
        // here would make every caller guard against a case that simply means "no files".
        ReviewSubmissionAssets Assets,

        // The files the PUBLISHED board references, which Assets is compared against. Supplied by
        // the server rather than derived here - see ParsePublishedFiles.
        IReadOnlyList<string> PublishedFiles,

        // Path -> SHA-256 for the published files the submission also carries, so byte-identical
        // pairs are dropped from the comparison. Empty when the server did not send it. See
        // ParsePublishedHashes.
        IReadOnlyDictionary<string, string> PublishedHashes,

        // Schematic name to the image file it is drawn from, so a moved highlight can be put back
        // on its own board. See ParseSchematicImages.
        IReadOnlyDictionary<string, string> SchematicImages,

        // Every submitted file with whose it is, whether a row uses it and what is published at
        // its path now - what ReviewFileComparison lists. Empty when the server did not send it.
        // See ParseSubmittedFiles.
        IReadOnlyList<SubmittedFileFact> SubmittedFiles,

        // Who must approve, who has, and what this account's approval would do - CRT.Data's
        // ApprovalStatus, as the server wrote it. Null from an older server; the window then
        // treats one approval as publishing, which is what that server did.
        ApprovalStatus? Approval = null,

        // The files publishing this would REMOVE from the BETA data (2026-09-25) - shown before
        // approving and sent back with the approval. Null from an older server.
        FileRemovalPreview? Removals = null,

        // Who last changed it in the Maintainer tab's table (2026-09-25), or null.
        ReviewAmendmentView? Amendment = null,

        // Who sent it and how their other submissions went (2026-09-26). Null from an older server.
        ReviewContributorFacts? Contributor = null);

    // A maintainer's change to a submission, as the submission view names it.
    public sealed record ReviewAmendmentView(int Version, string By, DateTimeOffset? AtUtc);

    // What saving a change in the table answered.
    public sealed record ReviewAmendResult(int Version, IReadOnlyList<ReviewFindingView> Warnings);

    // ###########################################################################################
    // What sending a change from a board's table answered (2026-10-03): the submission it became,
    // and any warnings its content raised. `Published`: it is in BETA at `Revision`, `RemovedFiles`
    // gone. Otherwise it was made but not published - `NotPublishedReason` says why - and waits under
    // Contributor Submissions.
    // ###########################################################################################
    public sealed record BoardEditResult(
        long SubmissionId,
        IReadOnlyList<ReviewFindingView> Warnings,
        bool Published = false,
        string? Revision = null,
        IReadOnlyList<string>? RemovedFiles = null,
        string? NotPublishedReason = null);

    // What removing unused files did, from the administrator's "Unused files" panel.
    public sealed record UnusedFileRemovalResult(
        string Tree,
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Kept,
        string? NotDoneBecause);

    // What rebuilding the checksum manifests did, from the administrator's entries on the Account screen
    // (2026-10-01). The headline and each tree's message are the SERVER's words, shown unchanged.
    public sealed record ManifestRebuildResult(
        string Headline,
        IReadOnlyList<ManifestRebuildTreeResult> Trees);

    public sealed record ManifestRebuildTreeResult(
        string Tree,
        bool Skipped,
        int Entries,
        string Message);

    public sealed record ReviewChangeSummaryView(
        bool IsNewBoard,
        IReadOnlyList<ReviewSectionView> Sections)
    {
        public int TotalChanges => this.Sections.Sum(section => section.TotalChanges);

        public bool HasChanges => this.TotalChanges > 0;
    }

    public sealed record ReviewSectionView(
        string Section,
        IReadOnlyList<string> Added,
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Changed,
        IReadOnlyList<ReviewRenameView> Renamed,
        IReadOnlyDictionary<string, IReadOnlyList<ReviewFieldChangeView>> FieldChanges)
    {
        public int TotalChanges =>
            this.Added.Count + this.Removed.Count + this.Changed.Count + this.Renamed.Count;
    }

    public sealed record ReviewRenameView(string From, string To, bool AlsoChanged);

    // ###########################################################################################
    // What a decision produced. Revision is empty for anything but an approval - only publishing
    // moves a board's revision.
    // ###########################################################################################
    // WaitingFor is filled when an approval was recorded but did not publish - the first of the
    // two a shared-file change needs (2026-09-25).
    // RemovedFiles: what a publishing approval removed from the BETA data because nothing used it.
    public sealed record ReviewDecisionResult(
        string State,
        string Revision,
        IReadOnlyList<ApproverRole>? WaitingFor = null,
        IReadOnlyList<string>? RemovedFiles = null);

    // ###########################################################################################
    // One field that moved, as the maintainer reads it.
    //
    // An empty Before means the field was BLANK and now has a value; an empty After means it was
    // CLEARED. Both are shown rather than rendered as nothing, because "(blank)" tells a maintainer
    // something and an empty cell tells them the app is broken.
    // ###########################################################################################
    public sealed record ReviewFieldChangeView(string Field, string Before, string After);

    // ###########################################################################################
    // A validation finding, as the maintainer reads it.
    //
    // IsError rather than a severity enum: there are exactly two values and the window's only
    // question is whether to colour it as a problem. An enum mirrored across the wire would be a
    // third place the vocabulary lives.
    // ###########################################################################################
    public sealed record ReviewFindingView(string Code, string Subject, string Message, bool IsError);

    // ###########################################################################################
    // One queue row as the app holds it.
    //
    // CreatedUtc is NULLABLE rather than defaulted, because "waiting since" is shown to the
    // maintainer and a missing timestamp rendered as 1970 would read as a submission that has been
    // waiting fifty years.
    // ###########################################################################################
    public sealed record ReviewQueueRow(
        long Id,
        string BoardId,
        string State,
        string Summary,
        string ContactEmail,
        DateTimeOffset? CreatedUtc,

        // Whether it adds or changes a shared file - which is why it is in an administrator's
        // queue rather than a maintainer's. Trailing with a default for an older server.
        bool TouchesSharedFiles = false,

        // Whether its board has no published board yet, and whether it waits for THIS account's
        // approval - CRT.Data's ReviewQueueEntry. Null when the server did not say.
        bool? IsNewBoard = null,
        bool? AwaitsYou = null,

        // When the contributor discarded their own draft of this board in CRT after sending it
        // (owner request, 2026-09-28) - CRT.Data's DraftDiscardContract. Null when they have not,
        // or from an older server.
        DateTimeOffset? DraftDiscardedUtc = null);
}
