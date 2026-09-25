using System;
using System.Linq;

namespace CRT.Maintainer.Handlers
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

        // ###########################################################################################
        // The DEFAULT server, and the only one anybody uses.
        //
        // *** NOT A PROMPT. *** The maintainer app talks to exactly one service and always will - there
        // is no second deployment for a maintainer to point at. Asking for it on the sign-in screen
        // made every launch start with a question whose answer never changes, and got it wrong the
        // moment somebody typed a trailing slash or forgot the scheme.
        //
        // It is still a CONSTANT rather than hardcoded at the call site so that a test, or a future
        // staging box, can override it in one place.
        //
        // *** IT CARRIES NO "/api" SUFFIX, and that is not the same as CRT's own constant. ***
        // Every method here appends the full path itself ("/api/review/queue"), so a base ending
        // in "/api" would produce "/api/api/review/queue" - a 404 on the very first request, from
        // an address that looks right in the source.
        //
        // CRT's AppConfig.CrtServerBaseUrl DOES include it, because SubmissionClient appends only
        // "/submissions". The two constants therefore differ on purpose. Caught by the route test
        // below rather than by a first launch against the live server.
        // ###########################################################################################
        public const string DefaultBaseAddress = "https://classic-repair-toolbox.dk";

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
        // The bytes CURRENTLY PUBLISHED for this submission's system, addressed by path (task 4).
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
        // Add and remove are both POSTs with a body: a system id carries slashes and cannot sit
        // in a route in front of the account id.
        // ###########################################################################################
        public static string AdminSystems(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/systems";

        public static string AdminAccounts(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/accounts";

        public static string AdminAddMaintainer(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/maintainers";

        public static string AdminRemoveMaintainer(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/maintainers/remove";

        // The "Unused files" screen (2026-09-25): the list for one tree ("beta" or "production"),
        // and removing the files the administrator chose.
        public static string AdminUnusedFiles(string baseAddress, string tree) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/unused-files?tree={Uri.EscapeDataString(tree ?? string.Empty)}";

        public static string AdminRemoveUnusedFiles(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/admin/unused-files/remove";

        // ###########################################################################################
        // BETA to production (2026-09-25) - ProductionEndpoints.MapProductionEndpoints. The plan and
        // the publish are POSTs with a body, because a system id carries slashes.
        // ###########################################################################################
        public static string ProductionList(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production";

        public static string ProductionPlan(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production/plan";

        public static string ProductionPublish(string baseAddress) =>
            $"{ReviewApiRoutes.Normalise(baseAddress)}/api/review/production/publish";

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
