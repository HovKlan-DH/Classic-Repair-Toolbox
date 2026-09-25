using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // How much ONE ADDRESS may submit (security review, 2026-09-25). Pure: it is handed what that
    // address has already submitted and the current time, and returns a verdict - the same shape
    // AuthRateLimitPolicy has for logins.
    //
    // *** CONTRIBUTING NEEDS NO ACCOUNT, SO THE ADDRESS IS ALL THERE IS TO LIMIT. *** Before this,
    // POST /api/submissions had no limit of any kind: each call stored a manifest of up to the
    // request size in the database and invited up to 5,000 files of 256 MB each. The abandoned-
    // upload sweep collected partial bytes after a day, and nothing else bounded anything.
    //
    // TWO BUCKETS, both over the same window:
    //
    //   COUNT - how many submissions. A real contributor sends a handful in a day: the work, then
    //           a correction after a maintainer asks for one. Twenty is generous for that and small
    //           for a script.
    //   BYTES - how much they asked to upload, as the server counted it at create (only the files
    //           it did not already hold). A whole new board is tens of megabytes; the largest
    //           shipped one is about 76 MB. Four gigabytes a day is far above any honest need and
    //           far below a full disk.
    //
    // *** AN ADDRESS IS WEAK, AND THIS IS NOT THE LAST LINE. *** One address can be a household or
    // a whole country behind CGNAT, and a determined sender can rotate addresses. That is why the
    // blob store also refuses work while the disk is below its reserve (BlobStore.HasRoomFor):
    // these limits stop the casual case, the reserve stops the determined one.
    //
    // A signed-in MAINTAINER or ADMINISTRATOR is not limited - they are trusted by the database, not
    // by a claim, and are the people most likely to submit many times while checking something.
    // An ordinary signed-in account is limited like anybody else, because anybody may register.
    // ###########################################################################################
    public static class SubmissionRateLimitPolicy
    {
        public const int MaxSubmissionsPerAddress = 20;

        public const long MaxUploadBytesPerAddress = 4L * 1024 * 1024 * 1024;

        public static readonly TimeSpan Window = TimeSpan.FromHours(24);

        // ###########################################################################################
        // May this address create another submission that asks to upload `bytesRequested`?
        //
        // `recent` is what the address submitted before - its entries' times and requested bytes,
        // in any order. Entries outside the window are ignored rather than trusting the caller to
        // have pruned them, the same rule AuthRateLimitPolicy follows.
        // ###########################################################################################
        public static RateLimitVerdict Check(
            IReadOnlyList<RecentSubmission> recent,
            long bytesRequested,
            DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(recent);

            DateTimeOffset cutoff = now - SubmissionRateLimitPolicy.Window;

            List<RecentSubmission> inWindow = recent.Where(entry => entry.CreatedUtc > cutoff).ToList();

            if (inWindow.Count >= SubmissionRateLimitPolicy.MaxSubmissionsPerAddress)
                return SubmissionRateLimitPolicy.Refused(inWindow, now);

            long bytesSoFar = inWindow.Sum(entry => Math.Max(0, entry.BytesRequested));

            // Checked as "would this one take it over", so a single upload larger than the whole
            // budget is refused too rather than slipping in because the bucket started empty.
            if (bytesSoFar + Math.Max(0, bytesRequested) > SubmissionRateLimitPolicy.MaxUploadBytesPerAddress)
                return SubmissionRateLimitPolicy.Refused(inWindow, now);

            return RateLimitVerdict.Allowed();
        }

        // Retry once the OLDEST entry in the window leaves it - the one whose expiry frees room.
        private static RateLimitVerdict Refused(List<RecentSubmission> inWindow, DateTimeOffset now)
        {
            if (inWindow.Count == 0)
                return RateLimitVerdict.Refused(SubmissionRateLimitPolicy.Window);

            TimeSpan retryAfter = inWindow.Min(entry => entry.CreatedUtc) + SubmissionRateLimitPolicy.Window - now;

            if (retryAfter < TimeSpan.Zero)
                retryAfter = TimeSpan.Zero;

            if (retryAfter > SubmissionRateLimitPolicy.Window)
                retryAfter = SubmissionRateLimitPolicy.Window;

            return RateLimitVerdict.Refused(retryAfter);
        }
    }

    // One earlier submission from an address: when, and how many bytes it asked to upload.
    public readonly record struct RecentSubmission(DateTimeOffset CreatedUtc, long BytesRequested);
}
