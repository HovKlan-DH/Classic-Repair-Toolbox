using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace CRT.Review.Handlers
{
    // ###########################################################################################
    // WHERE THE REVIEWER'S SIGNED-IN SESSION IS REMEMBERED between launches (maintainer request,
    // 2026-09-22: "only need to login once ... the password is complex and not something you want
    // to deal with for a low-volume thing like this").
    //
    // *** ReviewSession's HEADER USED TO FORBID THIS, AND ITS CONDITION IS MET RATHER THAN
    // IGNORED. *** It said: "Do not add remember-me without deciding where that file lives and who
    // can read it." Both answers are here - it lives beside CRT's own AppData files, and only the
    // logged-in Windows user can read it, because the token is DPAPI-encrypted to that user
    // (ReviewSessionProtection). A copy taken to another machine or another account is inert.
    //
    // *** ONLY THE TOKEN IS ENCRYPTED; THE REST IS PLAIN. *** Expiry, email and display name are
    // written as ordinary JSON so the file can be read by a person debugging it, and none of them
    // authorises anything. Encrypting the whole document would make a wrong stored expiry
    // impossible to diagnose without tooling, for no gain.
    //
    // *** THE STORED EXPIRY IS A HINT, NOT AN AUTHORITY. *** It exists so the app can skip a
    // request it knows will fail and show a sign-in screen instead. Whether the session is really
    // live is decided by the SERVER on the next call: a signed-out, revoked, expired or
    // locked-account session is refused there, and ReviewApiFailure.NotSignedIn makes this app
    // forget it. Trusting this file to say "still signed in" would be exactly the mistake.
    //
    // *** SLIDING EXPIRY IS WHAT MAKES THIS WORTH HAVING. *** The server pushes the expiry forward
    // as the app is used (SessionExtensionRules), so a regular user is never asked again, while a
    // machine that stops being used goes cold on its own. Note the app deliberately never calls
    // /api/accounts/refresh: that ROTATES the token, and a crash between the server poisoning the
    // old value and this file holding the new one would revoke every session for the account.
    //
    // FAILURES ARE SOFT THROUGHOUT, the same rule SubmissionReceiptStore follows: an unreadable or
    // unwritable file costs the convenience of staying signed in, never the ability to sign in.
    // ###########################################################################################
    internal static class ReviewSessionStore
    {
        public const string SessionFileName = "review-session.json";

        // The same AppData folder CRT itself uses, so a maintainer running both has one place to
        // look rather than two.
        public const string AppFolderName = "Classic-Repair-Toolbox";

        private static string _path = string.Empty;

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true
        };

        // ###########################################################################################
        // Resolves the real AppData location. Called once at startup.
        // ###########################################################################################
        public static void Initialise()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string directory = Path.Combine(appData, ReviewSessionStore.AppFolderName);

                Directory.CreateDirectory(directory);

                ReviewSessionStore.InitialiseAt(Path.Combine(directory, ReviewSessionStore.SessionFileName));
            }
            catch (Exception)
            {
                // No store, no remembered session. The app still works; it just asks for a password.
                ReviewSessionStore._path = string.Empty;
            }
        }

        // ###########################################################################################
        // Points the store at an explicit file - the test seam, mirroring UserSettings.LoadFrom and
        // SubmissionReceiptStore.LoadFrom. NEVER let a test call Initialise(), which would read and
        // write the real user's AppData folder.
        // ###########################################################################################
        internal static void InitialiseAt(string sessionFilePath) =>
            ReviewSessionStore._path = sessionFilePath ?? string.Empty;

        // ###########################################################################################
        // Whether staying signed in is possible on this machine at all.
        //
        // False off Windows, where there is no DPAPI equivalent and storing the token would mean
        // writing a publishing credential in plain text. The UI asks so it can tell the truth
        // instead of offering a tick box that silently does nothing.
        // ###########################################################################################
        public static bool IsSupported => ReviewSessionProtection.IsSupported;

        // ###########################################################################################
        // Remembers a session, or does nothing at all if it cannot be stored SAFELY.
        //
        // *** A FAILURE TO ENCRYPT MUST NEVER FALL BACK TO PLAINTEXT. *** Protect returns null off
        // Windows and on any DPAPI error, and the only correct response is to store nothing: a
        // token written in the clear is a publishing credential readable by anything running as
        // this user, which is precisely what the encryption was for.
        // ###########################################################################################
        public static void Remember(ReviewSession session)
        {
            if (session is null || string.IsNullOrWhiteSpace(ReviewSessionStore._path))
                return;

            byte[]? protectedToken = ReviewSessionProtection.Protect(session.BearerToken);

            if (protectedToken is null)
                return;

            try
            {
                var stored = new StoredSession
                {
                    ProtectedToken = Convert.ToBase64String(protectedToken),
                    ExpiresUtc = session.ExpiresUtc,
                    AccountId = session.AccountId,
                    Email = session.Email,
                    DisplayName = session.DisplayName
                };

                ReviewSessionStore.WriteAtomically(
                    JsonSerializer.Serialize(stored, ReviewSessionStore.JsonOptions));
            }
            catch (Exception)
            {
                // Losing the convenience is acceptable; failing the sign-in that just succeeded is
                // not.
            }
        }

        // ###########################################################################################
        // The remembered session, or NULL when there is none worth offering.
        //
        // *** AN EXPIRED STORED SESSION IS DISCARDED HERE, NOT RETURNED. *** Handing one back would
        // make the app open on the queue, fire a request, take a 401 and drop to the sign-in
        // screen - which reads as a bug. ReviewSession.IsUsableAt already applies the margin that
        // keeps a nearly-dead session from failing mid-request.
        //
        // Takes `now` rather than reading the clock, the same rule every flow in CRT.Server
        // follows, so expiry is testable without waiting.
        // ###########################################################################################
        public static ReviewSession? Recall(DateTimeOffset now)
        {
            if (string.IsNullOrWhiteSpace(ReviewSessionStore._path) || !File.Exists(ReviewSessionStore._path))
                return null;

            try
            {
                StoredSession? stored = JsonSerializer.Deserialize<StoredSession>(
                    File.ReadAllText(ReviewSessionStore._path));

                if (stored is null || string.IsNullOrWhiteSpace(stored.ProtectedToken))
                    return null;

                string? token = ReviewSessionProtection.Unprotect(
                    Convert.FromBase64String(stored.ProtectedToken));

                // Not ours, not this user's, or not this machine's. Clear it so the failure is not
                // retried on every launch forever.
                if (string.IsNullOrWhiteSpace(token))
                {
                    ReviewSessionStore.Forget();
                    return null;
                }

                var session = new ReviewSession(
                    token,
                    stored.ExpiresUtc,
                    stored.AccountId,
                    stored.Email ?? string.Empty,
                    stored.DisplayName ?? string.Empty);

                if (!session.IsUsableAt(now))
                {
                    ReviewSessionStore.Forget();
                    return null;
                }

                return session;
            }
            catch (Exception)
            {
                // Corrupt, truncated, hand-edited or written by a future version. Any of these
                // means "no usable stored session" - and leaving the file in place would repeat
                // the failure on every launch.
                ReviewSessionStore.Forget();
                return null;
            }
        }

        // ###########################################################################################
        // FORGETS the stored session - signing out, a rejected token, or an unreadable file.
        //
        // *** THE FILE IS DELETED, NOT BLANKED. *** An empty file left behind is indistinguishable
        // from a truncated write, and the next launch would have to decide which it was.
        // ###########################################################################################
        public static void Forget()
        {
            if (string.IsNullOrWhiteSpace(ReviewSessionStore._path))
                return;

            try
            {
                if (File.Exists(ReviewSessionStore._path))
                    File.Delete(ReviewSessionStore._path);
            }
            catch (Exception)
            {
                // Nothing useful to do. The session is already forgotten in memory, and the server
                // is the authority on whether the token still works.
            }
        }

        // ###########################################################################################
        // Writes via a temporary file and a replace, so an interrupted write cannot leave a
        // half-written document where a valid one was - the same reasoning AtomicJsonFile applies
        // in CRT.App.
        // ###########################################################################################
        private static void WriteAtomically(string json)
        {
            string temporary = ReviewSessionStore._path + ".tmp";

            File.WriteAllText(temporary, json);

            // Move with overwrite rather than Replace: Replace throws when the destination does
            // not exist, which is exactly the case on the very first sign-in.
            File.Move(temporary, ReviewSessionStore._path, overwrite: true);
        }

        // ###########################################################################################
        // The on-disk shape. Kept private so the file format is this class's business alone.
        // ###########################################################################################
        private sealed class StoredSession
        {
            [JsonPropertyName("protectedToken")]
            public string ProtectedToken { get; set; } = string.Empty;

            [JsonPropertyName("expiresUtc")]
            public DateTimeOffset ExpiresUtc { get; set; }

            [JsonPropertyName("accountId")]
            public long AccountId { get; set; }

            [JsonPropertyName("email")]
            public string? Email { get; set; }

            [JsonPropertyName("displayName")]
            public string? DisplayName { get; set; }
        }
    }
}
