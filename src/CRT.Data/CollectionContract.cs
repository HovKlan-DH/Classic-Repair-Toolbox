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

        // PUT - the owner stops sharing a collection, or shares it again (CollectionSharingRequest; server 15.1.0).
        public static string SharingOf(string id) => $"{CollectionRoutes.Collections}/{id}/sharing";
    }

    // ###########################################################################################
    // "STOP SHARING" (owner decision, 2026-11-27: "Offline, same link back"): PUT /api/collections/<id>/sharing with Shared
    // false takes the page offline - its link says it is not shared at the moment, its captures and comments are kept, its
    // deletion date keeps counting - and Shared true brings it back under the same link, as "Share" (the whole collection
    // sent again) does. Answered with the collection's CollectionAnswer; 410 once it is deleted.
    // ###########################################################################################
    public sealed record CollectionSharingRequest(bool Shared);

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
    //   DeleteAtUtc              when the server deletes it - the "Delete in" chosen, counted from the share (CollectionRules
    //                            .Lifetimes, at most a year ahead).
    //   Images                   the captures in the order the page shows them.
    //   Components               what CRT's component info window said of each component captured, as it was captured
    //                            (server 16.0.0) - the page shows it beside a capture seen large. Null from a CRT before it.
    // ###########################################################################################
    public sealed record CollectionContentRequest(
        string HardwareName,
        string BoardName,
        string? Region,
        string? Title,
        string? Description,
        IReadOnlyList<string>? ForumLinks,
        DateTimeOffset DeleteAtUtc,
        IReadOnlyList<CollectionImageEntry> Images,
        IReadOnlyList<CollectionComponent>? Components = null);

    // ###########################################################################################
    // One capture: which pin it is of and what the oscilloscope was set to (read back from it, as a baseline stores them).
    // `PinoutSha256`: CRT's drawing of the component's pinout as its component info window showed it, the capture's pin
    // marked - a picture uploaded as the captures are (server 16.0.0). (CRT's own baseline for the pin went with server 16.0.0
    // - owner decision, 2026-11-27: "that is not relevant".)
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
        string? Caption,
        DateTimeOffset? CapturedUtc,
        string? PinoutSha256 = null);

    // ###########################################################################################
    // A COMPONENT AS CRT'S COMPONENT INFO WINDOW SHOWED IT (owner request, 2026-11-27: "show something like the component info
    // window from the CRT application ... to get some context for this specific component and pin") - its title ("U7 | VIC-II |
    // 6569"), category line ("IC | 6569 | DIP-40"), short description, what it is there for (plain text), and its pins as the
    // window lists them, of the region shown. Words only, shown on the page as typed (encoded).
    // ###########################################################################################
    public sealed record CollectionComponent(
        string Label,
        string? Title,
        string? CategoryLine,
        string? Description,
        string? Purpose,
        IReadOnlyList<CollectionPin>? Pins);

    // One pin as the window lists it: its number or name, its signal name, direction, KiCad net and expected oscilloscope
    // reading (several lines when its baselines differ).
    public sealed record CollectionPin(string Pin, string? Name, string? Direction, string? Net, string? Reading);

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
