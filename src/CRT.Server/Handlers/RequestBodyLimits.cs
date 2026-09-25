using Handlers.DataHandling;
using Microsoft.AspNetCore.Http.Metadata;

namespace CRT.Server.Handlers
{
    // ###########################################################################################
    // How large a request body each route may send (security review, 2026-09-25). Program.cs applies
    // it once routing has chosen the endpoint and before that endpoint reads a byte.
    //
    // *** THE LIMIT WAS KESTREL'S DEFAULT, EVERYWHERE, AND NOBODY HAD CHOSEN IT. *** About 30 MB on
    // every route - including POST /api/submissions, which needs no account and stores the whole
    // manifest in the database as one LONGTEXT. So one anonymous request could park 30 MB of
    // whatever it liked in MariaDB, and a sign-in form accepted the same.
    //
    // THE NUMBERS COME FROM MEASUREMENT, not a guess. Every board shipped in Assets/Data was built
    // into a real manifest: the largest (C128 310378, 2,217 files) is about 1.2 MB of JSON. Eight
    // megabytes is several times that. A chunk may be twice the size CRT sends
    // (SubmissionFormat.UploadChunkBytes, shared with the client so the two cannot drift).
    // Everything else - a sign-in, a review comment of at most 4,000 characters - fits in 64 KB
    // with room to spare.
    //
    // *** THE MAINTAINER'S ROUTES THAT CARRY A BOARD OR A FILE LIST (2026-09-25). *** Saving the
    // maintainer's table sends the submission's whole rows - the same size as a manifest - so it gets
    // the manifest's limit; left at 64 KB, every save of a real board was refused with 413 before
    // the endpoint ran. Approving, publishing to production and removing unused files send back
    // the list of files the maintainer was shown for removal. The whole shipped tree listed that way
    // is about 600 KB (10,921 files), so two megabytes covers any such list.
    //
    // *** EACH LIMIT IS WRITTEN ON ITS ROUTE, where the route is mapped (code review, 2026-09-25):
    // `.WithBodyLimit(RequestBodyLimits.ManifestBytes)`. *** It used to be a second copy of the
    // route table here - URL segments and verbs matched by hand - so a new route that posts rows or
    // a file list, or a renamed segment, compiled, passed the tests of the routes this file knew,
    // and was refused with 413 in the maintainer's hands. That had already happened once (the table
    // save above). Now the limit sits beside the MapPost it governs, and RequestBodyLimitsTests
    // builds the server's real route table and fails on any route that reads a body without a
    // decision recorded there.
    //
    // DENY BY DEFAULT: an endpoint without a limit of its own - or no endpoint at all - gets the
    // SMALL one. A route that needs more has to say so, which is the moment somebody decides how
    // much more.
    // ###########################################################################################
    public static class RequestBodyLimits
    {
        public const long ManifestBytes = 8L * 1024 * 1024;

        public const long BlobChunkBytes = 2L * SubmissionFormat.UploadChunkBytes;

        public const long PathListBytes = 2L * 1024 * 1024;

        public const long DefaultBytes = 64L * 1024;

        // The limit for the endpoint routing chose, or the default when it chose none.
        public static long For(Endpoint? endpoint) =>
            endpoint?.Metadata.GetMetadata<BodyLimit>()?.Bytes ?? RequestBodyLimits.DefaultBytes;

        // Gives a route (or a group) a body limit other than the default.
        public static TBuilder WithBodyLimit<TBuilder>(this TBuilder builder, long bytes)
            where TBuilder : IEndpointConventionBuilder
        {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(bytes);

            return builder.WithMetadata(new BodyLimit(bytes));
        }
    }

    // ###########################################################################################
    // One route's limit, as endpoint metadata. It also implements ASP.NET Core's own
    // IRequestSizeLimitMetadata, so the framework's routing middleware - which applies that
    // interface where it finds it - sets the SAME number; Program.cs's middleware is what supplies
    // the default for every route without one.
    // ###########################################################################################
    public sealed record BodyLimit(long Bytes) : IRequestSizeLimitMetadata
    {
        long? IRequestSizeLimitMetadata.MaxRequestBodySize => this.Bytes;
    }
}
