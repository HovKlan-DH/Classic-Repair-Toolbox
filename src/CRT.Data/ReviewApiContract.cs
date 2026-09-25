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
    // *** WHY THIS FILE EXISTS. *** CRT.Maintainer calls CRT.Server over HTTP, so a renamed JSON field
    // compiles on both sides and fails in the maintainer's hands - CLAUDE.md's "One change, every side
    // of it" names it as the danger the compiler cannot see. The requests used to be records inside
    // the server's endpoint classes while the maintainer app sent ANONYMOUS objects with hand-typed
    // names, and every answer was an anonymous object on the server read back by hand-typed names in
    // ReviewApiParser. Renaming AmendRequest.ExpectedVersion would have made every amendment arrive
    // as version -1 and be refused as "changed since you opened it", with every test green.
    //
    //   - A REQUEST is now one record here, built by the maintainer app and bound by the server, so a
    //     rename moves both ends at once.
    //   - An ANSWER the maintainer app reads is one record here, written by the server. The maintainer app
    //     still reads it field by field (ReviewApiParser is forgiving on purpose), so
    //     CRT.Maintainer.Tests' ReviewWireContractTests serialises each record with the server's
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

        // The same settings as a ready object, for the maintainer app's requests and for tests.
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

    // ---- Requests: the maintainer application -> CRT.Server ----------------------------------------

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

    public sealed record UnusedFilesRemoveRequest(string? Tree, IReadOnlyList<string>? Files);

    // ---- Answers: CRT.Server -> the maintainer application ----------------------------------------

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
        FileRemovalPreview? Removals);

    // State is "published", or "awaiting" for the first of two approvals (then WaitingFor is set).
    public sealed record ProductionPublishAnswer(
        string SystemId,
        string State,
        string? Revision = null,
        int FilesCopied = 0,
        IReadOnlyList<ApproverRole>? WaitingFor = null,
        IReadOnlyList<string>? RemovedFiles = null);

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

    public sealed record ProductionListEntry(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? BetaRevision,
        string? BetaContentHash,
        string? ProductionRevision,
        DateTimeOffset? ProductionPublishedUtc);

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
