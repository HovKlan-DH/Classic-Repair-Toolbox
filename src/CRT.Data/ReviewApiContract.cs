using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE REVIEW API'S BODIES, IN ONE PLACE BOTH ENDS COMPILE AGAINST (code review, 2026-09-25).
    //
    // *** WHY THIS FILE EXISTS. *** CRT's Maintainer tab calls CRT.Server over HTTP, so a renamed JSON field
    // compiles on both sides and fails in the maintainer's hands - CLAUDE.md's "One change, every side
    // of it" names it as the danger the compiler cannot see. The requests used to be records inside
    // the server's endpoint classes while the Maintainer tab sent ANONYMOUS objects with hand-typed
    // names, and every answer was an anonymous object on the server read back by hand-typed names in
    // ReviewApiParser. Renaming AmendRequest.ExpectedVersion would have made every amendment arrive
    // as version -1 and be refused as "changed since you opened it", with every test green.
    //
    //   - A REQUEST is now one record here, built by the Maintainer tab and bound by the server, so a
    //     rename moves both ends at once.
    //   - An ANSWER the Maintainer tab reads is one record here, written by the server. The Maintainer tab
    //     still reads it field by field (ReviewApiParser is forgiving on purpose), so
    //     CRT.App.Tests' ReviewWireContractTests serialises each record with the server's
    //     settings and parses it with the real parser - a rename on either side fails there.
    //
    // WireSettings is the JSON the two agree on. The server applies it in Program.cs and the review
    // app serialises its requests with it, so a naming-policy change cannot reach one end only.
    //
    // ReviewTableData, FileRemovalPreview, ApprovalStatus, PromotionFile and UnusedFileListing are
    // shared records already and are used here as they are.
    // ###########################################################################################
    public static class ReviewApiContract
    {
        // ###########################################################################################
        // The server's settings on top of ASP.NET's own starting point (JsonSerializerDefaults.Web:
        // camelCase, case-insensitive reading). Nulls are left out - an absent field and a null one
        // read the same everywhere in ReviewApiParser.
        // ###########################################################################################
        public static void ApplyWireSettings(JsonSerializerOptions options)
        {
            options.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
            options.PropertyNameCaseInsensitive = true;
            options.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;
        }

        // The same settings as a ready object, for the Maintainer tab's requests and for tests.
        public static JsonSerializerOptions WireSettings { get; } = ReviewApiContract.CreateWireSettings();

        private static JsonSerializerOptions CreateWireSettings()
        {
            var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
            ReviewApiContract.ApplyWireSettings(options);

            // Frozen, since it is shared; freezing needs the resolver named (it is otherwise
            // chosen on first use, and MakeReadOnly throws without one).
            options.TypeInfoResolver = new DefaultJsonTypeInfoResolver();
            options.MakeReadOnly();
            return options;
        }
    }

    // ---- Requests: the Maintainer tab -> CRT.Server ----------------------------------------

    // Approve, reject, request changes. ExpectedRemovals (approve only): the files the maintainer was
    // shown the publish would remove - see ApprovePublishFlow step 5b. Null sends nothing.
    public sealed record ReviewDecisionRequest(string? Comment, IReadOnlyList<string>? ExpectedRemovals = null);

    // A maintainer's amendment: the rows as edited, and the amendment version the table opened at.
    public sealed record AmendRequest(int ExpectedVersion, SubmissionRows? Rows);

    public sealed record ProductionPlanRequest(string? SystemId);

    // ExpectedBetaContentHash: the BETA state the maintainer checked. ExpectedRemovals: the files they
    // were shown the promotion would remove from production.
    public sealed record ProductionPublishRequest(
        string? SystemId,
        string? ExpectedBetaContentHash,
        IReadOnlyList<string>? ExpectedRemovals = null);

    // Adding a maintainer to a system's pool, or removing one. The answer is not read beyond its status.
    public sealed record MaintainerChangeRequest(string? SystemId, long AccountId);

    // ###########################################################################################
    // INVITING A NEW MAINTAINER BY EMAIL (owner request, 2026-09-27: "I should be able to either
    // select an existing maintainer or invite a new maintainer via email"). The administrator names
    // an address with no account yet; the server mails it a code, and the person accepts it in CRT
    // Maintainer's sign-in window, which creates their account and puts them in the system's pool.
    //
    // POST /api/admin/maintainers/invite, /api/admin/maintainers/invitations/withdraw (administrator)
    // and POST /api/accounts/accept-invitation (anybody holding a code). The two admin answers are
    // { message }; the acceptance answers AcceptInvitationAnswer.
    // ###########################################################################################
    public sealed record MaintainerInviteRequest(string? SystemId, string? Email);

    public sealed record InvitationWithdrawRequest(long InvitationId);

    public sealed record AcceptInvitationRequest(string? Code, string? DisplayName, string? Password);

    public sealed record UnusedFilesRemoveRequest(string? Tree, IReadOnlyList<string>? Files);

    // ---- Answers: CRT.Server -> the Maintainer tab ----------------------------------------

    // ###########################################################################################
    // What a decision did. State is what the submission became ("merged", "approved", "rejected",
    // "changes_requested"). A publish adds its revision, content hash and the files it removed; the
    // first of two approvals adds who must still approve.
    // ###########################################################################################
    public sealed record ReviewDecisionAnswer(
        string State,
        string? Revision = null,
        string? ContentHash = null,
        IReadOnlyList<ApproverRole>? WaitingFor = null,
        IReadOnlyList<string>? RemovedFiles = null);

    // A saved amendment: its version, and any warnings it raised.
    public sealed record AmendAnswer(int Version, IReadOnlyList<ValidationFinding> Findings);

    // ###########################################################################################
    // What promoting one system from BETA to production would do. CanPublish is the one answer the
    // button follows, so the app cannot offer what the publish step would refuse.
    // ###########################################################################################
    public sealed record ProductionPlanAnswer(
        string SystemId,
        string? BetaRevision,
        string? BetaContentHash,
        bool IsAwaitingProduction,
        bool TouchesSharedFiles,
        bool CanPublish,
        string? Refusal,
        ApprovalStatus? Approval,
        int UnchangedCount,
        IReadOnlyList<PromotionFile> Files,
        IReadOnlyList<ValidationFinding> Problems,
        FileRemovalPreview? Removals,

        // ###########################################################################################
        // *** WHOSE WORK THIS PROMOTION CARRIES (owner request, 2026-09-27). *** The window used to
        // show the file copy list and nothing else - "I am not sure if the shown information in the
        // right-side panel is any helpful". A board is hundreds of files and a path says nothing
        // about whether the data is right, while the one thing a maintainer is actually about to do
        // - push named people's accepted work to every CRT user - was not on the screen at all.
        //
        // These are the submissions merged since the last promotion, which is exactly the set the
        // server already emails afterwards (ISubmissionStore.GetMergedSubmissionsAsync, called from
        // ProductionEndpoints.AfterPublishAsync). Shown BEFORE rather than only after.
        //
        // Optional, so an older maintainer build reads this answer unchanged.
        // ###########################################################################################
        IReadOnlyList<CarriedSubmission>? Carrying = null,

        // ###########################################################################################
        // *** THE FILE TREE'S OTHER HALF (owner request, 2026-09-28). *** The paths UnchangedCount
        // counts - with Files and Removals, every file of the system, so the Maintainer tab
        // can draw production after the publish (SystemFileEntries.ForPromotion). And where the
        // two trees are published, so a file can be opened from the tree exactly as CRT would
        // download it. All optional: an older maintainer build ignores them.
        // ###########################################################################################
        IReadOnlyList<string>? UnchangedFiles = null,
        string? BetaDataUrl = null,
        string? ProductionDataUrl = null);

    // ###########################################################################################
    // A SUBMISSION'S FILE TREE: the BETA data after approving it, against BETA now (owner request,
    // 2026-09-28) - GET /api/review/submissions/{id}/files, asked for when the maintainer opens it,
    // since working it out reads the board's folder and would slow every click on the queue.
    // `BetaDataUrl` is where BETA is published, for opening a file that is already there.
    // ###########################################################################################
    public sealed record SubmissionFilesAnswer(
        string SystemId,
        IReadOnlyList<SystemFileEntry> Files,
        string? BetaDataUrl = null);

    // ###########################################################################################
    // ROLLING A BETA BOARD BACK (owner decision, 2026-09-27) - the production window's "push back
    // to queue". Both routes POST because a system id carries slashes.
    //
    // The plan is shown before anything happens; the rollback carries the maintainer's REASON,
    // which is required - it is the contributor's only feedback.
    // ###########################################################################################
    public sealed record BetaRollbackPlanRequest(string SystemId);

    // `Reject` (owner request, 2026-09-28: "In the 'Beta > Prod' I would like a direct 'Reject'
    // button also - just like the normal queue. Then there is no need to push it back and then reject
    // it"): the same rollback, but the submissions it takes out of BETA are REJECTED with the comment
    // instead of going back to the queue. Optional, so an older CRT Maintainer pushes back as before.
    public sealed record BetaRollbackRequest(string SystemId, string Comment, bool Reject = false);

    // Kind is "restore" (production's bytes go back over BETA) or "remove" (a system never
    // promoted leaves the BETA tree). `Returning` is who goes back to the queue - named, because a
    // rollback reverts the whole board.
    //
    // `SharedRestored` (code review, 2026-09-27): the shared files the returning submissions
    // changed, put back to production's bytes. Every board citing a shared file sees it, which is
    // why the confirmation names them. Optional, so the plan still reads without them. A shared
    // file they ADDED is never removed by a rollback (BetaRollbackPlan's header), so there is no
    // list of removals - two fields that were always empty were dropped the same day.
    public sealed record BetaRollbackPlanAnswer(
        string SystemId,
        string Kind,
        IReadOnlyList<string> Restored,
        IReadOnlyList<string> Removed,
        IReadOnlyList<CarriedSubmission> Returning,
        IReadOnlyList<string>? SharedRestored = null);

    // `Rejected`: the submissions were rejected rather than returned to the queue - the request's
    // Reject, said back, so the Maintainer tab can tell an older server that ignored it (and pushed back)
    // from one that rejected. SubmissionsReturned counts them either way.
    public sealed record BetaRollbackAnswer(
        string SystemId,
        string Kind,
        int FilesRestored,
        int FilesRemoved,
        int SubmissionsReturned,
        bool Rejected = false);

    // ###########################################################################################
    // One merged submission a production promotion would carry out to everyone. `Comment` is the
    // contributor's own description - the same text the review queue lists it by.
    //
    // DraftDiscardedUtc (owner request, 2026-09-28): when the contributor discarded their own draft
    // of this board in CRT after sending it - see DraftDiscardContract. Null when they have not, and
    // from an older server.
    // ###########################################################################################
    public sealed record CarriedSubmission(
        long Id,
        string ContactEmail,
        string? Comment,
        DateTimeOffset? DecidedUtc,
        DateTimeOffset? DraftDiscardedUtc = null);

    // State is "published", or "awaiting" for the first of two approvals (then WaitingFor is set).
    public sealed record ProductionPublishAnswer(
        string SystemId,
        string State,
        string? Revision = null,
        int FilesCopied = 0,
        IReadOnlyList<ApproverRole>? WaitingFor = null,
        IReadOnlyList<string>? RemovedFiles = null);

    // ###########################################################################################
    // The review queue (2026-09-26 - an anonymous object until then, parsed by hand), oldest first.
    //
    // Beside the submission's own row, each entry carries what the queue list shows as badges
    // (owner request, 2026-09-26): whether its system has a published board at all (IsNewSystem -
    // the "no published board" answer the detail's comparison gives), and whether it waits for THIS
    // account's approval (AwaitsYou - ApprovalStatus.CanApprove, which the Approve button follows).
    // Null from a server that does not send them: the app shows no badge rather than a guess.
    //
    // *** NO UPLOAD TOKEN HASH, EVER. *** It is the contributor's capability for the submission;
    // echoing it to a maintainer would let them act as the contributor. The contact email is here
    // because it is the only way to reply - contributors have no account.
    // ###########################################################################################
    public sealed record ReviewQueueAnswer(
        bool CanPublish,
        bool IsAdministrator,
        int Count,
        IReadOnlyList<ReviewQueueEntry> Submissions);

    public sealed record ReviewQueueEntry(
        long Id,
        string SystemId,
        string State,
        string? Summary,
        string? ContactEmail,
        string? BaseRevision,
        DateTimeOffset? CreatedUtc,
        DateTimeOffset? DecidedUtc,
        bool TouchesSharedFiles,
        bool? IsNewSystem = null,
        bool? AwaitsYou = null,

        // When the contributor discarded their own draft of this board in CRT after sending it
        // (owner request, 2026-09-28) - see DraftDiscardContract. Null when they have not.
        DateTimeOffset? DraftDiscardedUtc = null);

    // ###########################################################################################
    // WHO SENT A SUBMISSION, AND HOW THEIR OTHER SUBMISSIONS WENT (owner request, 2026-09-26: "It
    // should be possible to see who it is (email) and how many contributions the contributor has
    // done, and some statics about accepted and rejected"). The submission detail's `contributor`.
    //
    // The same person as SubmissionReplacementRules means it: the same signed-in account, or the
    // same contact email (any case). Name is the account's display name, and null for the ordinary
    // contributor, who has no account.
    //
    // The counts are over the contributor's OTHER submissions, not this one:
    //   - Published: in the BETA or production data;
    //   - Waiting: for a decision, or for a second approval;
    //   - ChangesRequested / Rejected: sent back, or turned down, BY A MAINTAINER. A submission the
    //     automatic checks refused never reached anyone, and one replaced by the contributor's own
    //     newer submission is an earlier copy of the same work - neither is counted.
    // ###########################################################################################
    public sealed record ReviewContributorFacts(
        string? Email,
        string? Name,
        int Published,
        int Waiting,
        int ChangesRequested,
        int Rejected);

    public sealed record UnusedFilesRemoveAnswer(
        string Tree,
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Kept,
        string? NotDoneBecause);

    // ###########################################################################################
    // GET /api/review/production - the systems whose BETA is ahead of production, for this
    // account. Configured says whether the server can publish to production at all, so "nothing is
    // waiting" and "this server cannot do it" do not look the same. (Anonymous objects on the
    // server until the code review of 2026-09-25: a renamed BetaContentHash would have made every
    // publish to production send an empty hash and be refused as "changed in BETA".)
    // ###########################################################################################
    public sealed record ProductionListAnswer(bool Configured, IReadOnlyList<ProductionListEntry> Systems);

    // AwaitsYou (owner request, 2026-09-27 - the BETA button's badge counts "systems that need your
    // attention"): false when this account has already given its production approval for the BETA
    // state on offer, so the system waits for the OTHER approver, not for them. Optional, so an older
    // maintainer build reads this answer unchanged; null from an older server reads as "yours".
    public sealed record ProductionListEntry(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? BetaRevision,
        string? BetaContentHash,
        string? ProductionRevision,
        DateTimeOffset? ProductionPublishedUtc,
        bool? AwaitsYou = null,

        // Whether a submission this BETA state carries was sent by a contributor who has since
        // discarded their own draft in CRT (owner request, 2026-09-28) - see DraftDiscardContract.
        bool? CarriesDiscardedDraft = null);

    // ###########################################################################################
    // THE "SYSTEMS" SCREEN (owner request, 2026-09-27): every system, and one system's facts - who
    // maintains it, who has contributed to it and how that went, and its recent submissions.
    //
    // *** FOR EVERY MAINTAINER, NOT ONLY THE ADMINISTRATOR (owner decision, 2026-09-27: "Everything
    // for everyone"). *** Any account that may review anything sees every system's contributors and
    // their contact addresses, including systems it does not maintain. That is a deliberate widening
    // - the queue only ever showed the addresses of submissions the account could decide - and it is
    // the project owner's call, recorded in NewContributeStrategy.md.
    //
    // GET  /api/review/systems                     - SystemOverviewAnswer
    // POST /api/review/systems/detail  {systemId}  - SystemDetailAnswer (a POST: the id has slashes)
    // ###########################################################################################
    public sealed record SystemOverviewAnswer(IReadOnlyList<SystemOverviewEntry> Systems);

    // InBeta / InProduction: whether that data tree holds the board. InProduction is null when the
    // server has no production tree configured, so "not in production" is never claimed without
    // looking. The revisions are the `systems` row's, null for a shipped board nothing has published.
    // ViewsLast30Days: how often CRT users looked at the board in the last 30 days, BETA-source views
    // left out (BoardViewStatistics, 2026-09-27); null from a server older than that.
    public sealed record SystemOverviewEntry(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        bool InBeta,
        bool? InProduction,
        bool IsAwaitingProduction,
        bool IsAccepting,
        string? BetaRevision,
        string? ProductionRevision,
        DateTimeOffset? ProductionPublishedUtc,
        int MaintainerCount,
        int? ViewsLast30Days = null);

    public sealed record SystemDetailRequest(string? SystemId);

    // Invitations: the ones not accepted yet (2026-09-27) - sent to an ADMINISTRATOR only, the one
    // person who can invite or withdraw; null for everybody else. History: what has happened to the
    // system, newest first (SystemHistoryEntry, 2026-09-27); null from a server older than that.
    // Views: how often CRT users look at it (BoardViewStatistics, 2026-09-27); null from a server
    // older than that, or when the counts could not be read.
    public sealed record SystemDetailAnswer(
        SystemOverviewEntry System,
        IReadOnlyList<PoolMaintainerEntry> Maintainers,
        IReadOnlyList<SystemContributorEntry> Contributors,
        IReadOnlyList<SystemSubmissionEntry> Submissions,
        IReadOnlyList<MaintainerInvitationEntry>? Invitations = null,
        IReadOnlyList<SystemHistoryEntry>? History = null,
        BoardViewStatistics? Views = null);

    // An invitation to maintain a system that has not been accepted yet, withdrawn, or let expire.
    public sealed record MaintainerInvitationEntry(long Id, string Email, DateTimeOffset InvitedUtc, DateTimeOffset ExpiresUtc);

    // What accepting an invitation did: the address the account was made for (the app fills it into
    // the sign-in box), the systems it now maintains, and the sentence to show.
    public sealed record AcceptInvitationAnswer(string Email, IReadOnlyList<string> SystemIds, string Message);

    // ###########################################################################################
    // One contributor to ONE system, and how their submissions to it went - the same person as
    // ReviewContributorFacts and SubmissionReplacementRules mean it: the same ACCOUNT, or - among
    // submissions sent WITHOUT one - the same contact email. A signed-in submission and an anonymous
    // one are two contributors even under one address, on purpose (code review, 2026-09-27): an
    // anonymous sender's email is unverified, so merging them would let anybody add submissions to
    // an account holder's record. The same counting rules too: Accepted is merged (in BETA or
    // production), a rejection counts only when a maintainer made it, and a submission replaced by
    // the contributor's own newer one is not counted. Name is the account's display name, null for
    // the ordinary contributor with none.
    // ###########################################################################################
    public sealed record SystemContributorEntry(
        string? Email,
        string? Name,
        int Accepted,
        int Waiting,
        int ChangesRequested,
        int Rejected,
        DateTimeOffset? LastSubmittedUtc);

    // One submission to the system, newest first. State is the CONTRIBUTOR-FACING word
    // (ProductionPromotionRules.ContributorFacingState): "merged" is in BETA, "published" has reached
    // production, "returned" was pushed back out of BETA. DecisionComment is what the maintainer told
    // the contributor.
    //
    // DraftDiscardedUtc (owner request, 2026-09-28): when the contributor discarded their own draft
    // in CRT after sending this - see DraftDiscardContract.
    public sealed record SystemSubmissionEntry(
        long Id,
        string? ContactEmail,
        string? Summary,
        string State,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? DecidedUtc,
        string? DecisionComment,
        DateTimeOffset? DraftDiscardedUtc = null);

    // ###########################################################################################
    // WHERE A NEW SYSTEM GOES IN THE DROP-DOWN LISTS (owner request, 2026-09-27): "The maintainer
    // should order the new system, so it becomes visible in the right location for the drop-down
    // lists. This must be done before it can be pushed to BETA." Placed in the Systems screen, by
    // dragging it into the full list.
    //
    // GET  /api/review/systems/listing                        - SystemListingAnswer
    // POST /api/review/systems/listing  SetPlacementRequest   - SetPlacementAnswer
    //
    // The placement is one row of the newest main Excel data file: the two drop-down names, the
    // Overview notes, and which row it goes after (AfterExcelDataFile, that row's board workbook;
    // null puts it first). The server writes it into BETA's file when the system is published there
    // - or at once for a system already in BETA - and into production's, next to the same
    // neighbours, when it is promoted.
    // ###########################################################################################
    public sealed record SystemPlacement(string HardwareName, string BoardName, string Notes, string? AfterExcelDataFile);

    // The list as CRT shows it (BETA's newest main Excel data file, top to bottom), and every system
    // that is not in it yet but needs to be - in BETA already, or with a submission waiting.
    // HasList is false when BETA has no versioned main Excel data file to read.
    public sealed record SystemListingAnswer(
        bool HasList,
        IReadOnlyList<SystemListingRow> Rows,
        IReadOnlyList<UnlistedSystemEntry> Unlisted);

    public sealed record SystemListingRow(string SystemId, string HardwareName, string BoardName, string ExcelDataFile);

    // A system not in the list yet. Placement is what a maintainer saved, or null; Suggested is where
    // it would go untouched - at the end of the boards of the same hardware folder, under that
    // hardware's name, else at the end under its folder names. CanPlace is whether THIS account may
    // place it (a maintainer of the system, or the administrator).
    public sealed record UnlistedSystemEntry(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        bool InBeta,
        bool CanPlace,
        SystemPlacement? Placement,
        SystemPlacement Suggested);

    public sealed record SetPlacementRequest(
        string? SystemId,
        string? HardwareName,
        string? BoardName,
        string? Notes,
        string? AfterExcelDataFile);

    // ListedInBeta: the row was written into BETA's main Excel data file right away, because the
    // board is already there. Otherwise it is written when the system is published to BETA.
    public sealed record SetPlacementAnswer(SystemPlacement Placement, bool ListedInBeta, string Message);

    // GET /api/admin/systems - every system with its maintainers, for the administrator's Maintainers
    // window.
    public sealed record MaintainerSystemsAnswer(IReadOnlyList<MaintainerSystemEntry> Systems);

    public sealed record MaintainerSystemEntry(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? CurrentRevision,
        bool IsAccepting,
        IReadOnlyList<PoolMaintainerEntry> Maintainers);

    public sealed record PoolMaintainerEntry(long AccountId, string DisplayName, string Email);

    // ###########################################################################################
    // GET /api/admin/accounts - every account, to pick a maintainer from. *** NO PASSWORD HASH, no
    // session, no token *** - only what the administrator needs to recognise a person and see
    // whether they can be granted anything. A field added here goes to the administrator's screen.
    // ###########################################################################################
    public sealed record MaintainerAccountsAnswer(IReadOnlyList<MaintainerAccountEntry> Accounts);

    public sealed record MaintainerAccountEntry(
        long Id,
        string Email,
        string DisplayName,
        bool IsAdministrator,
        bool IsVerified,
        bool IsLocked);
}
