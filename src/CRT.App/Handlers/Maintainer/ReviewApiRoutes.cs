using System;
using System.Linq;
using CRT;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // The review API's routes, built in one place (NewContributeStrategy.md Phase 5, task 2).
    //
    // *** SEPARATE FROM THE HTTP CLIENT SO IT CAN BE TESTED. *** Building a URL is pure string
    // work, and it is the half of an API client that actually goes wrong: a missing slash, a
    // base address with a trailing one, an unescaped id. The socket half is an untested I/O
    // boundary by this project's own rules (CLAUDE.md test rule 6), so keeping the URLs out of it
    // means the part that can be wrong is the part that is covered.
    //
    // *** A TRAILING SLASH ON THE BASE ADDRESS MUST NOT CHANGE THE RESULT. *** It is configuration
    // a human types, so half of them will end in a slash. Uri's own combining rules would quietly
    // drop a path segment for one of the two spellings, which produces a 404 against a server
    // that is working perfectly - the kind of fault that gets blamed on the server for an hour.
    // ###########################################################################################
    public static class ReviewApiRoutes
    {
        public static string Queue(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/queue";

        public static string Submission(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/submissions/{submissionId}";

        public static string Login(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/accounts/login";

        // The deployed server's version (2026-10-04) - unauthenticated and public, so it is asked
        // with no session (ReviewApiClient.GetServerVersionAsync).
        public static string Health(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/health";

        // ###########################################################################################
        // The DEFAULT server, and the only one anybody uses.
        //
        // *** NOT A PROMPT. *** The Maintainer tab talks to exactly one service and always will - there
        // is no second deployment for a maintainer to point at. Asking for it on the sign-in screen
        // made every launch start with a question whose answer never changes, and got it wrong the
        // moment somebody typed a trailing slash or forgot the scheme.
        //
        // It is still a CONSTANT rather than hardcoded at the call site so that a test, or a future
        // staging box, can override it in one place.
        //
        // *** IT CARRIES NO "/api" SUFFIX, and that is not the same as AppConfig.CrtServerBaseUrl. ***
        // Every method here appends the full path itself ("/api/review/queue"), so a base ending
        // in "/api" would produce "/api/api/review/queue" - a 404 on the very first request, from
        // an address that looks right in the source.
        //
        // AppConfig.CrtServerBaseUrl DOES include it, because SubmissionClient appends only
        // "/submissions". Both are built from the one host, AppConfig.CrtServerRootUrl (since the
        // maintainer application became CRT's Maintainer tab, 2026-09-29), so they cannot point at
        // two different servers - but they still differ on purpose. Caught by the route test
        // rather than by a first launch against the live server.
        // ###########################################################################################
        public const string DefaultBaseAddress = AppConfig.CrtServerRootUrl;

        // Signing out. Revokes the ONE session presented, never the whole account - so signing out
        // on a laptop does not sign the project owner out at the bench.
        public static string Logout(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/accounts/logout";

        // Completing a reset: the code from the mail plus the new password.
        //
        // *** POST, AND THERE IS DELIBERATELY NO GET COUNTERPART. *** The mail used to carry a
        // link to "/api/accounts/reset", which the server maps nothing at - so every reset link
        // 404'd until 2026-09-22. It cannot be a link: setting a password needs the password.
        public static string ResetPassword(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/accounts/reset-password";

        public static string ForgotPassword(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/accounts/forgot-password";

        // Accepting an invitation to maintain a board (2026-09-27): the code from the mail, a name
        // and a password. Before anybody is signed in - it is how the account is made.
        public static string AcceptInvitation(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/accounts/accept-invitation";

        // ###########################################################################################
        // The signed-in maintainer's own account (2026-10-03) - the "Your account" window, and the
        // remembered sign-in refreshing its name and address. Server: AccountEndpoints, under its
        // "/api/accounts" group. A new address is two calls: ChangeEmail mails a code, and
        // ConfirmEmailChange spends it.
        // ###########################################################################################
        public static string Me(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/accounts/me";

        public static string ChangeName(string baseAddress) =>
            $"{ReviewApiRoutes.Me(baseAddress)}/name";

        public static string ChangeEmail(string baseAddress) =>
            $"{ReviewApiRoutes.Me(baseAddress)}/email";

        public static string ConfirmEmailChange(string baseAddress) =>
            $"{ReviewApiRoutes.Me(baseAddress)}/email/confirm";

        public static string ChangePassword(string baseAddress) =>
            $"{ReviewApiRoutes.Me(baseAddress)}/password";

        // ###########################################################################################
        // The three review DECISIONS (task 5).
        //
        // Separate routes rather than one taking the outcome as a body field, mirroring the
        // server: approving is irreversible and administrator-only while the other two are
        // neither, so a client bug that sent the wrong value would be a publish nobody asked for.
        // A wrong URL 404s; a wrong enum value publishes.
        // ###########################################################################################
        public static string Approve(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/approve";

        // The maintainer's table (2026-09-25): both boards to open it on, and saving a change.
        public static string SubmissionTable(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/table";

        public static string SubmissionAmend(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/amend";

        public static string Reject(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/reject";

        public static string RequestChanges(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/request-changes";

        // ###########################################################################################
        // The bytes a contributor UPLOADED, addressed by hash (task 4).
        //
        // The hash needs no escaping - it is 64 characters of lowercase hex or the server refuses
        // it - but it is escaped anyway, because "this particular value happens to be safe" is a
        // property of today's contract rather than of this method.
        // ###########################################################################################
        public static string SubmittedAsset(string baseAddress, long submissionId, string hash) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/submitted/{Uri.EscapeDataString(hash ?? string.Empty)}";

        // ###########################################################################################
        // The bytes CURRENTLY PUBLISHED for this submission's board, addressed by path (task 4).
        //
        // *** EACH SEGMENT IS ESCAPED, AND THE SLASHES BETWEEN THEM ARE NOT. *** Board data paths
        // routinely carry spaces ("Shared files/74LS08.png") and occasionally a "#", both of which
        // silently truncate or corrupt a raw URL - a "#" turns the rest of the path into a
        // fragment the server never sees, so the request arrives asking for a different file and
        // 404s. Escaping the WHOLE path instead would encode the separators too and the server's
        // catch-all would receive one long segment naming a file that does not exist.
        // ###########################################################################################
        public static string PublishedAsset(string baseAddress, long submissionId, string relativePath)
        {
            string escaped = string.Join('/',
                (relativePath ?? string.Empty)
                    .Split('/')
                    .Select(Uri.EscapeDataString));

            return $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/published/{escaped}";
        }

        // ###########################################################################################
        // The ADMINISTRATOR's routes (Phase 6 roles): the lists, and changing a pool. Under
        // "/api/admin" rather than "/api/review", matching AdminEndpoints' own group - see its
        // header for why the two are kept apart.
        //
        // Add and remove are both POSTs with a body: a board id carries slashes and cannot sit
        // in a route in front of the account id.
        // ###########################################################################################
        public static string AdminBoards(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/boards";

        public static string AdminAccounts(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/accounts";

        public static string AdminAddMaintainer(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/maintainers";

        public static string AdminRemoveMaintainer(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/maintainers/remove";

        // Inviting an address with no account yet, and taking an invitation back (2026-09-27).
        public static string AdminInviteMaintainer(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/maintainers/invite";

        public static string AdminWithdrawInvitation(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/maintainers/invitations/withdraw";

        // The "Unused files" screen (2026-09-25): the list for one tree ("beta" or "production"),
        // and removing the files the administrator chose.
        public static string AdminUnusedFiles(string baseAddress, string tree) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/unused-files?tree={Uri.EscapeDataString(tree ?? string.Empty)}";

        public static string AdminRemoveUnusedFiles(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/unused-files/remove";

        // Rebuilding both trees' dataChecksums.json by hand (2026-10-01). A POST with no body - the
        // one button does every configured tree, so there is nothing to choose.
        public static string AdminRebuildManifests(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/manifest/rebuild";

        // Deleting a board completely (2026-10-03): what it would remove, then the delete. Both
        // POST a body - a board id carries slashes.
        public static string AdminBoardDeletePlan(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/boards/delete/plan";

        public static string AdminBoardDelete(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/boards/delete";

        // The order of CRT's drop-down lists, written into BETA's and the stable source's main Excel
        // data files (2026-10-04) - every board BETA lists, in order, POSTed.
        public static string AdminBoardOrder(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/boards/order";

        // Resetting the contribution data for going live (2026-10-04): GET the counts, POST the reset
        // with their fingerprint - one path for both.
        public static string AdminDataReset(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/reset";

        // Which CRT versions call which route, over the last `days` days (2026-10-04).
        public static string AdminApiUsage(string baseAddress, int days) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/api-usage?days={days.ToString(System.Globalization.CultureInfo.InvariantCulture)}";

        // ###########################################################################################
        // BETA to production (2026-09-25) - ProductionEndpoints.MapProductionEndpoints. The plan and
        // the publish are POSTs with a body, because a board id carries slashes.
        // ###########################################################################################
        public static string ProductionList(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production";

        public static string ProductionPlan(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production/plan";

        // A submission's file tree: the BETA data after approving it (2026-09-28).
        public static string SubmissionFiles(string baseAddress, long submissionId) =>
            $"{ReviewApiRoutes.Submission(baseAddress, submissionId)}/files";

        // ###########################################################################################
        // A file in a PUBLISHED data tree, at the address CRT downloads it from (2026-09-28): the
        // tree's public base (the server's PublicDataBaseUrl or its production twin) and the
        // data-root-relative path, each segment escaped - a board folder or "Shared files" has
        // spaces. Null for a base that is not an absolute http(s) address, so nothing else is ever
        // asked for.
        // ###########################################################################################
        public static string? PublishedDataFile(string? dataBaseUrl, string relativePath)
        {
            if (!Uri.TryCreate(dataBaseUrl?.Trim(), UriKind.Absolute, out Uri? baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttps && baseUri.Scheme != Uri.UriSchemeHttp))
            {
                return null;
            }

            string escaped = string.Join('/',
                (relativePath ?? string.Empty)
                    .Replace('\\', '/')
                    .Split('/', StringSplitOptions.RemoveEmptyEntries)
                    .Select(Uri.EscapeDataString));

            return escaped.Length == 0 ? null : $"{baseUri.AbsoluteUri.TrimEnd('/')}/{escaped}";
        }

        public static string ProductionPublish(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production/publish";

        // Rolling BETA back and returning its submissions to the queue (2026-09-27).
        public static string BetaRollbackPlan(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production/rollback/plan";

        public static string BetaRollback(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production/rollback";

        // ###########################################################################################
        // The "Boards" screen (2026-09-27) - BoardEndpoints.MapBoardEndpoints. Every board, and
        // one board's facts; the detail is a POST because a board id carries slashes. Under
        // "/api/review" rather than "/api/admin": every maintainer may read it.
        // ###########################################################################################
        public static string Boards(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards";

        public static string BoardDetail(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards/detail";

        // Where a new board goes in the drop-down lists (2026-09-27): GET the list, POST a placement.
        public static string BoardListing(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards/listing";

        // ###########################################################################################
        // A board's Board data and Files views (2026-10-03) - BoardEndpoints' four POSTs: BETA's
        // board, what a table edit would remove from BETA, the edit itself (published straight to
        // BETA), and every file the board uses.
        // ###########################################################################################
        public static string BoardDataTable(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards/table";

        public static string BoardEditCheck(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards/edit/check";

        public static string BoardEdit(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards/edit";

        public static string BoardFiles(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/boards/files";

        // ###########################################################################################
        // The base address with any trailing slashes removed.
        //
        // Throws on a blank one rather than building a relative URL that would be attempted
        // against whatever the client's own base happened to be - a request going somewhere
        // unintended is worse than one that does not go at all, and this is a maintainer's
        // credentials being posted.
        // ###########################################################################################
        private static string Normalise(string baseAddress)
        {
            if (string.IsNullOrWhiteSpace(baseAddress))
                throw new ArgumentException("A server address is required.", nameof(baseAddress));

            return baseAddress.Trim().TrimEnd('/');
        }
    }
}
