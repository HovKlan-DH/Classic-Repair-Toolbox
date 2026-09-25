using System.Globalization;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Where a blob lives on disk, and how a partial upload is addressed. PURE string work - it
    // builds paths and never touches the filesystem, so every rule here is a unit test.
    //
    // CONTENT-ADDRESSED: a blob's name IS its SHA-256. Two systems carrying the same schematic
    // image store one file, which is what makes "shared images already present under another board
    // are never re-sent" true rather than aspirational. It also means a blob never needs updating:
    // the same hash is always the same bytes, so a write either creates a file or is redundant.
    //
    // FANNED OUT TWO LEVELS by the first four hex characters: ab/cd/abcd1234... A flat directory
    // of tens of thousands of files is slow to enumerate on every filesystem and unpleasant to
    // inspect by hand; 256 x 256 buckets keeps any one directory small for far more blobs than
    // this project will ever hold.
    //
    // WHY THE HASH IS RE-VALIDATED HERE rather than trusted from the manifest: this function turns
    // a caller-supplied string into a filesystem path. If it were ever called with an unvalidated
    // value, "../../etc/passwd" would become a write target. Validating at the point of use costs
    // nothing and means the safety does not depend on every caller having remembered.
    // ###########################################################################################
    public static class BlobStorePaths
    {
        // Partial uploads live beside the finished store rather than inside it, so a directory walk
        // of the blob store never has to distinguish complete files from in-flight ones - and a
        // crash mid-upload cannot leave something that looks like a valid blob.
        public const string PartialFolderName = "partial";

        public const string BlobFolderName = "blobs";

        // ###########################################################################################
        // The final resting place of a completed blob, relative to the blob store root.
        //
        // Returns false for a malformed hash rather than throwing: the value may have come from a
        // request, and the caller turns this into a rejection.
        // ###########################################################################################
        public static bool TryGetBlobPath(string blobRoot, string hash, out string path)
        {
            path = string.Empty;

            if (string.IsNullOrWhiteSpace(blobRoot))
                return false;

            if (!SubmissionPathRules.IsValidHash(hash))
                return false;

            // Safe to index: IsValidHash has established 64 characters of lowercase hex, so there
            // is no path separator, no "..", and nothing that could escape the root.
            string first = hash[..2];
            string second = hash[2..4];

            path = Path.Combine(blobRoot, BlobStorePaths.BlobFolderName, first, second, hash);

            return true;
        }

        // ###########################################################################################
        // Where an in-progress upload accumulates.
        //
        // SCOPED TO THE SUBMISSION, not just to the hash. Two contributors uploading the same new
        // file at the same time would otherwise append into one partial file and produce garbage
        // that matches neither. Scoping costs a duplicate upload in that rare case and removes an
        // interleaving bug that would be almost impossible to reproduce.
        //
        // The submission id is a server-issued integer, never a caller-supplied string, so it
        // cannot carry a traversal - but it is formatted invariantly so a server whose locale uses
        // different digits still produces the path the next request will look for.
        // ###########################################################################################
        public static bool TryGetPartialPath(string blobRoot, long submissionId, string hash, out string path)
        {
            path = string.Empty;

            if (string.IsNullOrWhiteSpace(blobRoot))
                return false;

            if (submissionId <= 0)
                return false;

            if (!SubmissionPathRules.IsValidHash(hash))
                return false;

            path = Path.Combine(
                blobRoot,
                BlobStorePaths.PartialFolderName,
                submissionId.ToString(CultureInfo.InvariantCulture),
                hash);

            return true;
        }

        // ###########################################################################################
        // The folder holding every partial upload for one submission, so finalising or abandoning
        // a submission can clear them in one go.
        //
        // Blobs from abandoned submissions must be garbage-collected or the disk fills quietly -
        // named as a trap in NewContributeStrategy.md. Having one folder per submission is what
        // makes that a directory delete rather than a query.
        // ###########################################################################################
        public static bool TryGetPartialFolder(string blobRoot, long submissionId, out string path)
        {
            path = string.Empty;

            if (string.IsNullOrWhiteSpace(blobRoot) || submissionId <= 0)
                return false;

            path = Path.Combine(
                blobRoot,
                BlobStorePaths.PartialFolderName,
                submissionId.ToString(CultureInfo.InvariantCulture));

            return true;
        }
    }
}
