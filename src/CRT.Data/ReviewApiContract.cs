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
    // What adding or removing a maintainer answers: the request's two ids echoed back (2026-10-04 - an
    // anonymous object with these same fields until then). The Maintainer tab reads nothing of it.
    public sealed record MaintainerChangeAnswer(string? SystemId, long? AccountId);

    public sealed record MaintainerInviteRequest(string? SystemId, string? Email);

    public sealed record InvitationWithdrawRequest(long InvitationId);

    public sealed record AcceptInvitationRequest(string? Code, string? DisplayName, string? Password);

    // ###########################################################################################
    // A MAINTAINER'S OWN ACCOUNT (owner request, 2026-10-03: the maintainer "can edit his/her own
    // email address and name"). The Maintainer tab's "Your account" window, signed in:
    //
    //   GET  /api/accounts/me                                           - AccountAnswer
    //   POST /api/accounts/me/name           ChangeNameRequest          - AccountChangeAnswer
    //   POST /api/accounts/me/email          ChangeEmailRequest         - AccountChangeAnswer
    //   POST /api/accounts/me/email/confirm  ConfirmEmailChangeRequest  - AccountChangeAnswer
    //   POST /api/accounts/me/password       ChangePasswordRequest      - AccountChangeAnswer
    //
    // A refusal is a 400 carrying { message }, or { errors: [...] } for a name or password that
    // breaks a rule - the two shapes the password reset already answers. 429 when too many mails
    // were asked for from one address.
    //
    // THE SESSION IS ENOUGH - no current password (owner decision, 2026-10-03: "as I see it as you
    // are already logged in"). A NEW ADDRESS still needs a code mailed to it, which proves the new
    // mailbox is the maintainer's; only the capitals changing needs no code - it is the same mailbox.
    // ###########################################################################################
    public sealed record ChangeNameRequest(string? DisplayName);

    public sealed record ChangeEmailRequest(string? NewEmail);

    public sealed record ConfirmEmailChangeRequest(string? Code);

    public sealed record ChangePasswordRequest(string? NewPassword);

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

    // An amendment or a system's table edit that its checks refused: the sentence to show, and the
    // findings behind it (2026-10-04 - an anonymous object with these same fields until then).
    public sealed record FindingsRefusalAnswer(string Error, IReadOnlyList<ValidationFinding> Findings);

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
        string? ProductionDataUrl = null,

        // ###########################################################################################
        // Every file's size, by path (owner request, 2026-10-04: sizes in every file tree) - for the
        // files Files, Removals and UnchangedFiles name: a removed one as production holds it, every
        // other as BETA does, which is what it will be. The tree is drawn by the Maintainer tab from
        // the plan (SystemFileEntries.ForPromotion), so the sizes travel beside it. A file missing
        // from it shows no size. Optional: an older maintainer build ignores it.
        // ###########################################################################################
        IReadOnlyDictionary<string, long>? FileSizes = null);

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
    //
    // *** THE CONTRIBUTOR'S WHOLE RECORD (owner request, 2026-09-30: "all information about the
    // contributor, to get an honest opinion if this person can be trusted, so all data we can see
    // for whatever he/she has contributed, and which has been accepted and rejected etc."). *** The
    // Maintainer tab's Contributor view. Optional and trailing, so an older Maintainer tab reads
    // this answer unchanged and an older server's answer reads as "not said":
    //   - SignedIn: whether THIS submission was sent from an account (a verified address) or
    //     without one (an address anybody could have typed);
    //   - AccountCreatedUtc: since when that account exists - null without one;
    //   - Submissions: the OTHER submissions the counts count, newest first - each one's system,
    //     description, state in CRT's words, dates and what the contributor was told.
    //
    // *** PublishedToStable (owner request, 2026-10-01: "[1] published to stable") *** - how many
    // of Published have reached the STABLE source; the rest of Published is in BETA only. Null
    // from a server older than 3.9.0, which the Maintainer tab then words as plain "published".
    // ###########################################################################################
    public sealed record ReviewContributorFacts(
        string? Email,
        string? Name,
        int Published,
        int Waiting,
        int ChangesRequested,
        int Rejected,
        bool? SignedIn = null,
        DateTimeOffset? AccountCreatedUtc = null,
        IReadOnlyList<ContributorSubmissionEntry>? Submissions = null,
        int? PublishedToStable = null);

    // ###########################################################################################
    // One of a contributor's other submissions, in the Contributor view (2026-09-30). State is the
    // CONTRIBUTOR-FACING word, as SystemSubmissionEntry's is (ProductionPromotionRules
    // .ContributorFacingState): "merged" is in BETA, "published" has reached production, "returned"
    // was pushed back out of BETA - so a submission reads the same on every screen. DecisionComment
    // is what the contributor was told.
    // ###########################################################################################
    public sealed record ContributorSubmissionEntry(
        long Id,
        string SystemId,
        string? Summary,
        string State,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? DecidedUtc,
        string? DecisionComment);

    public sealed record UnusedFilesRemoveAnswer(
        string Tree,
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Kept,
        string? NotDoneBecause);

    // ###########################################################################################
    // POST /api/admin/manifest/rebuild - rebuilding dataChecksums.json for both data trees by hand
    // (owner request, 2026-10-01). No request body: there is nothing to choose, the button does
    // both trees. See CRT.Server's ManifestRebuildFlow for why it exists.
    //
    // `Headline` and each entry's `Message` are the SERVER's words, shown unchanged - so a server
    // that learns a new outcome (a third tree, a new reason to skip one) says so without a CRT
    // release. `Entries` is the number of files listed, or -1 for a failure; `Skipped` is a tree
    // this server has not configured, which is not a failure.
    // ###########################################################################################
    public sealed record ManifestRebuildAnswer(
        string Headline,
        IReadOnlyList<ManifestRebuildEntry> Trees);

    public sealed record ManifestRebuildEntry(
        string Tree,
        bool Skipped,
        int Entries,
        string Message);

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
        bool? CarriesDiscardedDraft = null,

        // ###########################################################################################
        // Whether it waits for the ADMINISTRATOR because only administrators publish to stable and
        // this account is not one (2026-10-05, StablePublishing). AwaitsYou is then false, which
        // alone reads as "you approved, it is with the other approver" - wrong words for a system
        // the account never approved (code review, 2026-10-05). Optional: null from an older server.
        // ###########################################################################################
        bool? WaitsForAdministrator = null);

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
    // BetaContentHash: the `systems` row's record of BETA's content, which every publish to BETA and
    // every push-back moves (2026-10-04) - so the Systems screen can tell its open table is out of
    // date without reading the board. Null for a board nothing has published, or from an older server.
    // ListedInBeta / ListedInStable (owner request, 2026-10-04: "I do not expect there should be cases
    // where something can only be listed in stable? If so, it must be flagged in the left-sided menu
    // 'Systems' list"): whether that source's newest main Excel data file lists the system in CRT's
    // drop-down lists. Null when that list could not be read - or, for the stable one, when the server
    // has no stable source - so a missing row is never claimed without looking.
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
        int? ViewsLast30Days = null,
        string? BetaContentHash = null,
        bool? ListedInBeta = null,
        bool? ListedInStable = null);

    // `Tree` (2026-10-04) chooses the data a system's Board data and Files are read from:
    // DataTreeNames.Production for the stable source, anything else - or none, as before - BETA.
    // The detail route ignores it.
    public sealed record SystemDetailRequest(string? SystemId, string? Tree = null);

    // Invitations: the ones not accepted yet (2026-09-27) - sent to an ADMINISTRATOR only, the one
    // person who can invite or withdraw; null for everybody else. History: what has happened to the
    // system, newest first (SystemHistoryEntry, 2026-09-27); null from a server older than that.
    // Views: how often CRT users look at it (BoardViewStatistics, 2026-09-27); null from a server
    // older than that, or when the counts could not be read.
    //
    // AddressesHidden (owner request, 2026-10-05): true when the account does not maintain the
    // system, so the server sent NO email address - not its maintainers' (an empty Email), its
    // contributors' or its submissions' (null), nor any in its history (names instead, where there is
    // an account). The Maintainer tab says so rather than showing a contributor without an account
    // as one who gave no address. False from a server older than that, which sent every address.
    public sealed record SystemDetailAnswer(
        SystemOverviewEntry System,
        IReadOnlyList<PoolMaintainerEntry> Maintainers,
        IReadOnlyList<SystemContributorEntry> Contributors,
        IReadOnlyList<SystemSubmissionEntry> Submissions,
        IReadOnlyList<MaintainerInvitationEntry>? Invitations = null,
        IReadOnlyList<SystemHistoryEntry>? History = null,
        BoardViewStatistics? Views = null,
        bool AddressesHidden = false);

    // ###########################################################################################
    // A SYSTEM'S BOARD DATA AND FILES ON THE "SYSTEMS" SCREEN (owner request, 2026-10-03: "all the
    // same functionalities, as the 'Contributor Submissions' has ... Board data (and I should be
    // able to do the same edits)", and "Files (should not show changed files - just list all
    // files)"). Each route POSTs SystemDetailRequest, because a system id carries slashes.
    //
    // POST /api/review/systems/table       {systemId} - SystemTableAnswer: BETA's board as rows.
    // POST /api/review/systems/edit/check  SystemEditRequest - SystemEditCheckAnswer: what the edit
    //                                      would remove from BETA, asked before the reason is.
    // POST /api/review/systems/edit        SystemEditRequest - SystemEditAnswer: the edit, PUBLISHED.
    // POST /api/review/systems/files       {systemId} - SystemFilesAnswer: every file the system uses.
    //
    // *** AN EDIT GOES STRAIGHT TO BETA (owner decision, 2026-10-03: "I do not think it should be
    // necessary for that change to first go to 'Contributor Submissions' queue - instead it should go
    // directly to the next queue, 'BETA > Stable', so it can directly be tested in BETA"). *** It
    // replaced the same day's "Becomes a submission". The server still builds a submission from BETA's
    // own files, from the maintainer's account, with the reason as its description - and then
    // approves it at once as that maintainer, through the ordinary approval, so every rule an approval
    // keeps (the files it removes, shared files, one submission in BETA per system) still holds.
    //
    // *** NOT WHILE THE SYSTEM WAITS IN BETA > STABLE (owner decision, same day: "If a system is
    // already in 'BETA > Stable' queue, then it should simply disallow it, even if this is coming
    // from a maintainer"). *** `MayEdit` is then false, with the reason.
    //
    // `ExpectedRemovals` is the list the check answered and the maintainer was shown; the edit is
    // refused (409) when the publish would now remove anything else - an approval's own rule.
    //
    // `Fingerprint` is BETA's board as the table was opened on it, sent back with the edit: BETA
    // published again meanwhile would otherwise be silently reverted by a table read before it.
    // `MayEdit` is false for a system this account does not maintain (anybody may READ every system -
    // "Everything for everyone"), with the reason in `MayNotEditReason`.
    // ###########################################################################################
    public sealed record SystemTableAnswer(
        string SystemId,
        string Fingerprint,
        SubmissionRows Rows,
        bool MayEdit,
        string? MayNotEditReason = null,
        string? BetaDataUrl = null,

        // Where the STABLE source is published (2026-10-04) - set, with BetaDataUrl null, when the
        // table is the stable source's (SystemDetailRequest.Tree), which is never editable.
        string? ProductionDataUrl = null);

    // `Summary` is the reason the maintainer gave - the submission's description. The check ignores it.
    public sealed record SystemEditRequest(
        string? SystemId,
        string? Fingerprint,
        string? Summary,
        SubmissionRows? Rows,
        IReadOnlyList<string>? ExpectedRemovals = null);

    // What publishing the edit would remove from BETA - files only this board used, which it stops
    // citing. Empty for nearly every edit.
    public sealed record SystemEditCheckAnswer(IReadOnlyList<string> Removals);

    // ###########################################################################################
    // What sending the edit did. `Published`: it is in BETA at `Revision`, with `RemovedFiles` gone,
    // and the system waits under BETA > Stable. Not published: the submission was made but the
    // publish did not happen (`NotPublishedReason` - something changed between the check and the
    // publish, say), so it waits under Contributor Submissions like any other, nothing lost.
    // `Findings` are the warnings its content raised (errors refuse it instead).
    // ###########################################################################################
    public sealed record SystemEditAnswer(
        long SubmissionId,
        IReadOnlyList<ValidationFinding> Findings,
        bool Published = false,
        string? Revision = null,
        IReadOnlyList<string>? RemovedFiles = null,
        string? NotPublishedReason = null);

    public sealed record SystemFilesAnswer(
        string SystemId,
        IReadOnlyList<SystemFileEntry> Files,
        string? BetaDataUrl = null,

        // Where the stable source is published (2026-10-04) - set when the files are the stable
        // source's, whose entries open from there (SystemFileSource.Production).
        string? ProductionDataUrl = null);

    // An invitation to maintain a system that has not been accepted yet, withdrawn, or let expire.
    public sealed record MaintainerInvitationEntry(long Id, string Email, DateTimeOffset InvitedUtc, DateTimeOffset ExpiresUtc);

    // What accepting an invitation did: the address the account was made for (the app fills it into
    // the sign-in box), the systems it now maintains, and the sentence to show.
    public sealed record AcceptInvitationAnswer(string Email, IReadOnlyList<string> SystemIds, string Message);

    // ###########################################################################################
    // The signed-in account, as GET /api/accounts/me answers it (2026-10-03 - an anonymous object
    // with these same fields until then, read by nothing). MaintainerOf: the systems whose pools it
    // is in - for an administrator, who reviews every system anyway, only the systems they were
    // named a maintainer of (2026-10-05), often none. *** NO PASSWORD
    // HASH, NO SESSION, NO TOKEN. ***
    // ###########################################################################################
    public sealed record AccountAnswer(
        long Id,
        string Email,
        string DisplayName,
        bool IsVerified,
        bool IsAdministrator,
        IReadOnlyList<string> MaintainerOf,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? LastLoginUtc = null);

    // What a change to the account did: the sentence to show, and the account as it is now. While a
    // new address waits for its code, Account is null and CodeSent is true - the address has NOT
    // changed yet.
    public sealed record AccountChangeAnswer(string Message, AccountAnswer? Account = null, bool CodeSent = false);

    // ###########################################################################################
    // A session, as signing in and refreshing answer it (2026-10-04 - an anonymous object with these
    // same fields until then, so nothing held the server's names to ReviewApiParser.ParseLogin's).
    // *** RefreshToken IS THE BEARER TOKEN *** - see ReviewSession's header for that naming trap.
    // ###########################################################################################
    public sealed record SessionAnswer(string RefreshToken, DateTimeOffset ExpiresUtc, SessionAccountAnswer Account);

    public sealed record SessionAccountAnswer(long Id, string Email, string DisplayName, bool IsVerified);

    // ###########################################################################################
    // One submission's detail, as GET /api/review/submissions/{id} answers it and
    // ReviewApiParser.ParseSubmission reads it (2026-10-04 - an anonymous object with these same
    // fields until then, so the API compatibility check could not see it). Changes is null when the
    // submission's payload could not be loaded, which a maintainer must still be able to open;
    // Amendment is null until a maintainer has changed the submission in the table.
    // ###########################################################################################
    public sealed record SubmissionDetailAnswer(
        bool CanPublish,
        ApprovalStatus Approval,
        ReviewQueueEntry Submission,
        SubmissionManifest? Manifest,
        ReviewContributorFacts Contributor,
        IReadOnlyList<ValidationFinding> Findings,
        ReviewChangeSummary? Changes,
        IReadOnlyList<string> PublishedFiles,
        IReadOnlyDictionary<string, string> PublishedHashes,
        IReadOnlyDictionary<string, string> SchematicImages,
        IReadOnlyList<SubmittedFileFact> SubmittedFiles,
        FileRemovalPreview Removals,
        SubmissionAmendmentFact? Amendment);

    // Who last changed a submission in the Maintainer tab's table, and to which version.
    public sealed record SubmissionAmendmentFact(int Version, string By, DateTimeOffset AtUtc);

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
    //
    // Changes (owner request, 2026-10-04): what it changed as it went into BETA - SubmissionChanges,
    // recorded by the publish. Null for one never published, or published before they were kept.
    public sealed record SystemSubmissionEntry(
        long Id,
        string? ContactEmail,
        string? Summary,
        string State,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? DecidedUtc,
        string? DecisionComment,
        DateTimeOffset? DraftDiscardedUtc = null,
        SubmissionChanges? Changes = null);

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

    // ###########################################################################################
    // THE ORDER OF THE DROP-DOWN LISTS, SET BY THE ADMINISTRATOR (owner request, 2026-10-04: "a
    // possibility to be able to sort the list of systems, which then gets saved to both sources
    // (BETA + stable) after my save"). POST /api/admin/systems/order.
    //
    // The request is every system BETA's list holds, by id, in the order wanted - all of them, so
    // a list that changed since it was read (a system placed meanwhile) is refused, not half
    // applied. The stable source's list follows the same order (MasterListing.ArrangeAs).
    //
    // The answer: whether each list was rewritten - StableChanged null when the server has no stable
    // source - and, when BETA's list was saved but something after it was not (the stable list, a
    // checksum manifest), what. CRT then shows that as well as the save.
    // ###########################################################################################
    public sealed record SystemOrderRequest(IReadOnlyList<string>? SystemIds);

    public sealed record SystemOrderAnswer(bool BetaChanged, bool? StableChanged, string? Problem = null);

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

    // ###########################################################################################
    // DELETING A SYSTEM COMPLETELY (owner request, 2026-10-03: "as admin, I should be able to have
    // a possibility to delete a system completely, which then will remove it from everywhere -
    // including BETA and stable sources ... part of the 'Admin' menu ... with a confirmation box").
    // Administrator only - CRT.Server's SystemDeletionFlow.
    //
    // POST /api/admin/systems/delete/plan  SystemDetailRequest - SystemDeletePlanAnswer: what would go.
    // POST /api/admin/systems/delete       SystemDeleteRequest - SystemDeleteAnswer: what went.
    //
    // *** WHAT WAS CONFIRMED IS WHAT IS DELETED. *** Fingerprint covers every file in the system's
    // folder in both trees, its rows in both drop-down lists and its submissions. The delete sends it
    // back and is refused (409) when the server would now delete something else - a publish or a new
    // submission in between - so nothing goes that the administrator was not shown.
    //
    // BlockedBecause: why it cannot be deleted at all, in the server's words (an older main Excel
    // data file lists it, or another board uses a file in its folder). Null when it can.
    //
    // OpenSubmissions: those still in play - waiting for review, approved once, or in BETA. Their
    // contributors are mailed (owner decision, 2026-10-03: "Delete them and mail the contributors"),
    // so Reason is REQUIRED when there are any: it is quoted in that mail. State is the word the
    // contributor is told (SubmissionReceiptPresenter.DescribeState reads it).
    // ###########################################################################################
    public sealed record SystemDeleteRequest(string? SystemId, string? Fingerprint, string? Reason = null);

    public sealed record SystemDeletePlanAnswer(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string Fingerprint,
        int BetaFiles,
        int ProductionFiles,
        bool ListedInBeta,
        bool ListedInProduction,
        bool HasRecord,
        int Submissions,
        int Maintainers,
        int Invitations,
        IReadOnlyList<SystemDeleteOpenSubmission> OpenSubmissions,
        string? BlockedBecause = null);

    public sealed record SystemDeleteOpenSubmission(
        long Id,
        string State,
        string Contributor,
        string? Summary,
        DateTimeOffset CreatedUtc);

    public sealed record SystemDeleteAnswer(
        string SystemId,
        int BetaFilesRemoved,
        int ProductionFilesRemoved,
        int SubmissionsDeleted,
        int ContributorsMailed);

    // ###########################################################################################
    // RESETTING THE CONTRIBUTION DATA (owner request, 2026-10-04: "When I go-live with this, it
    // should not have old data visible ... the real sources of BETA and stable must not be touched,
    // but all contributor and maintainer data should go away"). Administrator only - CRT.Server's
    // DataResetFlow - and only while the server's AllowDataReset setting is on.
    //
    // GET  /api/admin/reset  - DataResetPlanAnswer: what a reset would delete, and whether it may.
    // POST /api/admin/reset  DataResetRequest - DataResetAnswer: what was deleted.
    //
    // *** WHAT WAS SHOWN IS WHAT IS DELETED. *** Fingerprint covers the submissions, accounts,
    // maintainers, invitations and system records; the reset sends it back and is refused (409) when
    // any of them changed since - a new submission, a new account. The history, the board views and
    // the API usage counts are counted but NOT in it: they grow by themselves, every few minutes, and
    // would make every reset refuse.
    //
    // Accounts counts the accounts that go - every one but the administrators. Administrators counts
    // those kept, with their sessions: somebody must be able to sign in afterwards, and the
    // administrator pressing the button stays signed in.
    //
    // Systems counts the system RECORDS that go - all of them. What a record holds about the data
    // trees (BETA's and the stable source's content hashes, a new system's place in the drop-down
    // lists) is re-made by the next publish; every flow already treats a system with no record as
    // one the pipeline has never touched, which after a reset is what every system is.
    //
    // NotEnabledBecause: the server's words for why the reset is switched off; null when it is on.
    // ###########################################################################################
    public sealed record DataResetRequest(string? Fingerprint);

    public sealed record DataResetPlanAnswer(
        bool IsEnabled,
        string Fingerprint,
        int Submissions,
        int Accounts,
        int Administrators,
        int Maintainers,
        int Invitations,
        int Systems,
        int HistoryEntries,
        int BoardViews,
        int ApiUsageRows,
        string? NotEnabledBecause = null);

    public sealed record DataResetAnswer(
        int SubmissionsDeleted,
        int AccountsDeleted,
        int MaintainersDeleted,
        int InvitationsDeleted,
        int SystemsDeleted,
        int HistoryEntriesDeleted,
        int BoardViewsDeleted,
        int ApiUsageRowsDeleted,
        int StoredFilesRemoved);

    // ###########################################################################################
    // WHICH CRT VERSIONS CALL WHICH ROUTE (owner request, 2026-10-04: "how about tracking the API
    // end-points, to see if it is possible to retire any, if almost no versions uses it any more").
    // Administrator only - CRT.Server's ApiUsageFlow.
    //
    // GET /api/admin/api-usage?days=90 - ApiUsageAnswer.
    //
    // Routes: EVERY route the server maps, called or not - a route nobody called is the answer most
    // worth seeing. Method and Route are the route's pattern ("POST", "/api/review/systems/edit"),
    // never a real path, so no id is ever counted. Area is ClientVersionPolicy's ("Forever",
    // "Submissions", "Maintainer"), sent as text so a new area never breaks an older reader;
    // NeverRetired is true for the routes every CRT ever released sends (owner decision, 2026-10-04:
    // the check-in, feedback, board views, the health check and the 2.x contribution address).
    //
    // Versions: per CRT version that called it in the window - the version a User-Agent names, or
    // ApiUsageVersion.NotCrt for a request naming none (a browser, an uptime check).
    //
    // Installations: per CRT version, how many installations launched it in the window (distinct
    // addresses in the launch check-ins) and how many launches - so "who still runs 3.0.0" can be
    // read beside "who still calls this route". The addresses never leave the server.
    // ###########################################################################################
    public sealed record ApiUsageAnswer(
        int Days,
        IReadOnlyList<ApiUsageRoute> Routes,
        IReadOnlyList<ApiUsageInstallations> Installations);

    public sealed record ApiUsageRoute(
        string Method,
        string Route,
        string Area,
        bool NeverRetired,
        long Calls,
        DateTimeOffset? LastUtc,
        IReadOnlyList<ApiUsageVersion> Versions);

    public sealed record ApiUsageVersion(string Version, long Calls, DateTimeOffset LastUtc)
    {
        // A request whose User-Agent names no CRT version.
        public const string NotCrt = "(not CRT)";

        // More different versions than the server counts apart in a day - see ApiUsageCounter.
        public const string Other = "(other)";
    }

    public sealed record ApiUsageInstallations(string Version, int Installations, int Launches);

    // ###########################################################################################
    // What GET /api/health answers with - {"status":"ok","version":"4.3.2","utc":"..."}. Built by
    // CRT.Server's HealthReport; read by the Maintainer tab, which shows the version under
    // Account > "Server version" (owner request, 2026-10-04: "I would like to see the server version
    // listed, so it is clear to me what has been deployed"). Here, not in CRT.Server, since 2026-10-04
    // so both ends use the one record (ReviewWireContractTests).
    //
    // ApiRevision (server 4.6.0) is the server's ClientVersionContract.ApiRevision, which the tab shows
    // beside its own under Account > "Server version" - so a maintainer sees whether CRT or the
    // server is behind. Null from an older server.
    // ###########################################################################################
    public sealed record HealthStatus(string Status, string Version, DateTimeOffset Utc, int? ApiRevision = null);
}
