using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE COLLECTION ROUTES' BODIES (2026-11-26; CollectionRules says what a collection is). One record per request and answer,
    // built by CRT and bound by the server - or written by the server and read by CRT - in ReviewApiContract's WireSettings, so
    // a renamed field fails the build or CollectionWireContractTests (CRT.App.Tests), never a person's share.
    //
    // HOW CRT SHARES A COLLECTION, every request signed in (Authorization: Bearer):
    //
    //   1. POST /api/collections (or PUT /api/collections/<id> for one shared before) with the WHOLE collection as it is now -
    //      its captures named by checksum. The server answers 201/200, or 409 naming the captures it does not hold yet
    //      (CollectionRefusalAnswer.MissingImages).
    //   2. Each of those PUT /api/collections/images/<sha256>, the PNG's bytes as the body - checked by CollectionPng.
    //   3. The first request again.
    //
    // So a collection changed later sends only the captures added since, and a capture the server lost (the account's upload
    // of it gone after its collection went) is simply sent again.
    // ###########################################################################################
    public static class CollectionRoutes
    {
        public const string Collections = "/api/collections";

        public const string Images = "/api/collections/images";

        public const string Mine = "/api/collections/mine";

        // The header a first share (POST) carries the collection's local id in - CollectionRules.IsClientKey. The same first
        // share sent again under it, its answer lost the first time, changes the page it made (code review, 2026-11-27).
        public const string ClientKeyHeader = "Idempotency-Key";

        // The public pages, under the site (Apache forwards "/collections/" to the service, as it does "/api/").
        public const string PagePath = "/collections";

        // A collection's comments (step 3, 2026-11-26): listed and written by CRT for its owner.
        public static string CommentsOf(string id) => $"{CollectionRoutes.Collections}/{id}/comments";

        public static string CommentOf(string id, long commentId) => $"{CollectionRoutes.CommentsOf(id)}/{commentId}";
    }

    // ###########################################################################################
    // COMMENTS ON A COLLECTION'S CAPTURES (step 3, 2026-11-26; CollectionComments says the rules) - written on the page by
    // anybody signed in, and by the owner from CRT too:
    //
    //   GET    /api/collections/<id>/comments              CollectionCommentsAnswer - the owner's (hidden ones too)
    //   POST   /api/collections/<id>/comments              CollectionCommentRequest - a comment on one capture
    //   PUT    /api/collections/<id>/comments/<cid>        CollectionCommentEditRequest - its writer's
    //   DELETE /api/collections/<id>/comments/<cid>        its writer's, or an administrator's
    //   PUT    /api/collections/<id>/comments/<cid>/hidden CollectionCommentHideRequest - the collection's owner's
    // ###########################################################################################
    public sealed record CollectionCommentRequest(string ImageSha256, string Body);

    public sealed record CollectionCommentEditRequest(string Body);

    public sealed record CollectionCommentHideRequest(bool Hidden);

    // One comment as the owner's CRT reads it. `ByOwner`: written by the collection's owner; `Mine`: by the account asking.
    // `AuthorLocked`: its writer's account is locked, so nobody but the administrator sees it - only ever true for the
    // administrator, the one owner who is answered such a comment (code review, 2026-11-27).
    public sealed record CollectionCommentAnswer(
        long Id,
        string ImageSha256,
        string AuthorName,
        bool ByOwner,
        bool ByAdministrator,
        bool Mine,
        string Body,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? EditedUtc,
        bool Hidden,
        bool AuthorLocked = false);

    public sealed record CollectionCommentsAnswer(IReadOnlyList<CollectionCommentAnswer> Comments);

    // ###########################################################################################
    // A collection as CRT shares it - the whole of it, every time.
    //
    //   HardwareName, BoardName  the board, as CRT's drop-down lists name it; Region the one shown when captured, if the board
    //                            has regions.
    //   Title, Description       the owner's words, shown on the page as typed (encoded, never read as markup).
    //   ForumLinks               the threads where it was posted (CollectionRules.IsForumLink).
    //   DeleteAtUtc              when the server deletes it (CollectionRules.Lifetimes, at most 5 years ahead).
    //   Images                   the captures in the order the page shows them.
    // ###########################################################################################
    public sealed record CollectionContentRequest(
        string HardwareName,
        string BoardName,
        string? Region,
        string? Title,
        string? Description,
        IReadOnlyList<string>? ForumLinks,
        DateTimeOffset DeleteAtUtc,
        IReadOnlyList<CollectionImageEntry> Images);

    // ###########################################################################################
    // One capture: which pin it is of, what the oscilloscope was set to (read back from it, as a baseline stores them), and
    // CRT's own baseline for that pin, if the board has one - the page shows the two side by side.
    // ###########################################################################################
    public sealed record CollectionImageEntry(
        string Sha256,
        string ComponentLabel,
        string? Pin,
        string? SignalName,
        string? ExpectedReading,
        string? TimePerDivision,
        string? VoltsPerDivision,
        string? TriggerLevel,
        CollectionBaseline? Baseline,
        string? Caption,
        bool LooksWrong,
        DateTimeOffset? CapturedUtc);

    // CRT's baseline for the pin: its picture by checksum and board file name (the published tree stores it under both -
    // DataTreeLayout.StoredFilePath), and how it was measured. Null checksum: the board has a baseline row with no picture.
    public sealed record CollectionBaseline(
        string? Sha256,
        string? FileName,
        string? TimePerDivision,
        string? VoltsPerDivision,
        string? TriggerLevel);

    // A collection taken or changed.
    public sealed record CollectionAnswer(
        string Id,
        string Url,
        DateTimeOffset CreatedUtc,
        DateTimeOffset DeleteAtUtc,
        int ImageCount);

    // A refusal: what was wrong, and - with 409 - the captures (by checksum) to upload before sending the collection again.
    public sealed record CollectionRefusalAnswer(
        IReadOnlyList<string> Errors,
        IReadOnlyList<string>? MissingImages = null);

    // A capture taken (PUT /api/collections/images/<sha256>).
    public sealed record CollectionImageUploadAnswer(
        string Sha256,
        int Width,
        int Height);

    // GET /api/collections/mine - the account's collections still shown, newest first.
    public sealed record MyCollectionsAnswer(IReadOnlyList<CollectionSummaryAnswer> Collections);

    // `CommentCount` and `LatestCommentId`: the comments shown on its page (none hidden or deleted) - CRT marks a collection
    // with comments newer than it has seen.
    public sealed record CollectionSummaryAnswer(
        string Id,
        string Url,
        string? Title,
        string HardwareName,
        string BoardName,
        string? Region,
        int ImageCount,
        DateTimeOffset CreatedUtc,
        DateTimeOffset UpdatedUtc,
        DateTimeOffset DeleteAtUtc,
        int CommentCount = 0,
        long LatestCommentId = 0);
}
