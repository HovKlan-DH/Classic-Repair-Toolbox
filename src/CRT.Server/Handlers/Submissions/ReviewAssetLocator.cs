using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHICH bytes a maintainer may fetch, and what they are served as
    // (NewContributeStrategy.md Phase 5, task 4).
    //
    // Task 4's remaining sub-parts - a moved highlight drawn on the schematic, images side by
    // side, scope baselines plotted - all need IMAGE BYTES, from both sides of the comparison.
    // This decides what may be handed over. ReviewAssetEndpoints is the rim that sends it.
    //
    // *** PURE, AND SPLIT OUT FOR EXACTLY THAT REASON. *** This is a file-disclosure boundary: a
    // path arriving in a request that reaches the data tree hands a maintainer any file the service
    // can read, including appsettings.Production.json. That decision is unit tested here rather
    // than verified by reading an endpoint.
    //
    // *** THE TWO SIDES ARE GUARDED DIFFERENTLY, BECAUSE THEY TAKE DIFFERENT INPUT. ***
    //
    //   SUBMITTED bytes are addressed by HASH. A 64-character lowercase hex string cannot carry a
    //   traversal, so containment is not the risk here. The risk is CROSS-SUBMISSION READING: the
    //   blob store is shared and content-addressed, so without a scope check "may review
    //   submission 7" would mean "may read any blob anyone has ever uploaded", given a hash - and
    //   hashes travel in manifests, which maintainers see. So the hash must be one the named
    //   submission's own manifest references.
    //
    //   PUBLISHED bytes are addressed by a caller-supplied PATH into the data tree. That IS the
    //   traversal risk, and it goes through SubmissionPathRules - the same containment the write
    //   paths use, resolved against this board's own folder rather than against the tree, so a
    //   maintainer opening a C64 submission cannot read an Amstrad board through it.
    //
    // *** THE IDENTITY IS UNTRUSTED TOO. *** Manufacturer/Hardware/Board arrive inside the
    // submission, so the board folder is resolved through the same rules before anything is
    // contained to it - otherwise a manufacturer of ".." would move the folder being contained to
    // and every subsequent check would guard the wrong place. Same reasoning as
    // PublishedBoardLocator, which has the longer write-up.
    // ###########################################################################################
    public static class ReviewAssetLocator
    {
        // ###########################################################################################
        // The ONLY content types these bytes are ever served as.
        //
        // *** AN ALLOWLIST, AND IT MUST STAY ONE. *** These files are contributor-supplied and are
        // served from the server's own origin. A file served as text/html would run script in that
        // origin against a signed-in maintainer's session, and a type inferred from the extension
        // lets the attacker pick the type by picking the name. Everything not on this list is
        // octet-stream, which downloads instead of rendering - which also happens to be the right
        // behaviour for the datasheet a maintainer wants to open.
        //
        // SVG is deliberately ABSENT despite being an image: it is an XML document that can carry
        // script, and browsers execute it when it is navigated to directly. No shipped board uses
        // one.
        // ###########################################################################################
        private static readonly Dictionary<string, string> ContentTypesByExtension =
            new(StringComparer.OrdinalIgnoreCase)
            {
                [".png"] = "image/png",
                [".jpg"] = "image/jpeg",
                [".jpeg"] = "image/jpeg",
                [".gif"] = "image/gif",
                [".bmp"] = "image/bmp",
                [".webp"] = "image/webp"
            };

        public const string DefaultContentType = "application/octet-stream";

        // ###########################################################################################
        // May this review fetch this submitted blob?
        //
        // Scoped to the manifest rather than merely checking the store holds it - see the header.
        // A null manifest allows nothing: a submission whose payload could not be loaded must not
        // fall open.
        // ###########################################################################################
        public static bool IsSubmittedBlobAllowed(SubmissionManifest? manifest, string? hash)
        {
            if (manifest is null)
                return false;

            // Re-validated here as well as in BlobStorePaths. Belt and braces on a value that
            // becomes a filesystem path, so the safety does not depend on the store having
            // remembered - the same reasoning BlobStorePaths' own header gives.
            if (!SubmissionPathRules.IsValidHash(hash))
                return false;

            // Ordinal: the store keys on the exact string, so accepting another spelling would
            // permit a fetch the store then misses, which reads as "the file is gone" rather than
            // "you asked wrongly".
            return manifest.Files.Any(file =>
                string.Equals(file.Sha256, hash, StringComparison.Ordinal));
        }

        // ###########################################################################################
        // Where a PUBLISHED file for this board is, if the request may have it at all.
        //
        // Returns false for anything unsafe, anything the submission does not reference, and
        // anything that is not there - the caller turns all three into one 404. They are
        // deliberately indistinguishable to the client: letting a caller tell "refused" from
        // "absent" lets the tree be probed for what exists.
        //
        // ###########################################################################################
        // *** THE PATH IS RESOLVED AGAINST THE DATA ROOT, NOT THE BOARD FOLDER (fixed
        // 2026-09-23). ***
        //
        // A board stores its file references RELATIVE TO THE DATA TREE ROOT - "Commodore/C64/250407/
        // Board Layout 250407 NTSC.png" - which is how the desktop app resolves them too
        // (Main.BoardSelection.cs combines DataManager.DataRoot with the stored value). Resolving
        // against the board folder therefore looked for
        // "<root>/Commodore/C64/250407/Commodore/C64/250407/Board Layout 250407 NTSC.png",
        // which never exists, and every published image answered 404. The maintainer saw "No
        // published file at this path" beside a picture that is in fact published.
        //
        // It also could not have worked for SHARED files: "Commodore/Shared files/Component
        // images/6526.png" lives outside the board folder by design, so a board-relative base
        // excludes an entire legitimate category of board reference.
        //
        // *** CONTAINMENT IS NOT WEAKENED, AND THE SCOPE IS NOW STRICTER THAN IT WAS. *** The
        // resolve still refuses anything escaping the root, and the request must additionally name
        // a file THIS SUBMISSION'S OWN BOARD REFERENCES (the check below). The board folder was
        // serving as a crude scope limit; naming the referenced files is the real one, and unlike
        // the folder it cannot be satisfied by an unrelated file that happens to sit nearby.
        // ###########################################################################################
        //
        // *** OR A FILE THE PUBLISHED BOARD CITES (owner request, 2026-09-26). *** A row changed to
        // another picture, or deleted, leaves its OLD file cited by the published board alone - and
        // the old picture is exactly the "before" side of the comparison, in the change summary and
        // in the table's hover card. Scoped to this submission's own board still: the published
        // board is the one the submission would replace. `publishedBoardFiles` is asked only when
        // the submission does not cite the path itself, so the ordinary request never reads it.
        // ###########################################################################################
        public static bool TryLocatePublishedFile(
            string? dataTreeRoot,
            SubmissionManifest? manifest,
            string? relativePath,
            out string resolvedPath,
            Func<IReadOnlyCollection<string>>? publishedBoardFiles = null)
        {
            resolvedPath = string.Empty;

            if (manifest is null)
                return false;

            if (string.IsNullOrWhiteSpace(dataTreeRoot))
                return false;

            if (string.IsNullOrWhiteSpace(relativePath))
                return false;

            // THE SCOPE. Only a file this submission's board - or the published board it would
            // replace - actually references may be fetched, so the route cannot be used to read
            // arbitrary files out of the data tree.
            if (!ReviewAssetLocator.IsReferencedByBoard(manifest, relativePath) &&
                !ReviewAssetLocator.IsListed(publishedBoardFiles?.Invoke(), relativePath))
            {
                return false;
            }

            // THE CONTAINMENT. Resolved against the data tree root, which is the base the stored
            // paths are written against; escaping the root is still refused.
            if (!SubmissionPathRules.TryResolve(dataTreeRoot, relativePath, out string candidate, out _))
                return false;

            if (!File.Exists(candidate))
                return false;

            resolvedPath = candidate;
            return true;
        }

        // ###########################################################################################
        // Whether the submission's own board references this exact path.
        //
        // Ordinal, like every other path comparison against this tree: the server's filesystem is
        // Linux and the data tree is case-sensitive, so folding case here would admit a path the
        // board does not actually name.
        //
        // The submitted rows are used rather than the published board because they are what the
        // manifest carries, and the two agree on every file the comparison offers - the Maintainer tab
        // only ever asks for a path that appeared in one of the two lists it was given.
        // ###########################################################################################
        private static bool IsReferencedByBoard(SubmissionManifest manifest, string relativePath)
        {
            string wanted = relativePath.Replace('\\', '/').Trim();

            if (wanted.Length == 0)
                return false;

            foreach (string candidate in ReviewAssetLocator.BoardFilePaths(manifest))
            {
                if (string.Equals(candidate, wanted, StringComparison.Ordinal))
                    return true;
            }

            return false;
        }

        // Whether a path is in a list of board file paths. Ordinal, for the reason IsReferencedByBoard
        // gives.
        private static bool IsListed(IReadOnlyCollection<string>? paths, string relativePath)
        {
            string wanted = relativePath.Replace('\\', '/').Trim();

            return wanted.Length > 0 &&
                   paths is not null &&
                   paths.Any(candidate => string.Equals(candidate, wanted, StringComparison.Ordinal));
        }

        // Every file path the submitted rows name, through the SHARED collector - the same one the
        // submission and the published-file list both use, so a new file source cannot become
        // unfetchable here while working everywhere else.
        private static IEnumerable<string> BoardFilePaths(SubmissionManifest manifest)
        {
            // Built against no published board, so the result is the submission's OWN rows -
            // exactly the set the Maintainer tab was handed and can ask about.
            BoardData board = PublishMerge.Build(manifest, published: null);

            return SubmissionManifestBuilder.CollectReferencedFiles(board);
        }

        // ###########################################################################################
        // The board's own folder inside the data tree, resolved through the same rules that guard
        // every other use of these three untrusted values.
        // ###########################################################################################
        private static bool TryResolveBoardFolder(
            string dataTreeRoot,
            SubmissionManifest manifest,
            out string boardFolder)
        {
            boardFolder = string.Empty;

            string relative = string.Join('/',
                new[] { manifest.Manufacturer, manifest.Hardware, manifest.Board }
                    .Where(part => !string.IsNullOrWhiteSpace(part)));

            if (string.IsNullOrWhiteSpace(relative))
                return false;

            if (!SubmissionPathRules.TryResolve(dataTreeRoot, relative, out string resolved, out _))
                return false;

            boardFolder = resolved;
            return true;
        }

        // ###########################################################################################
        // What to serve these bytes as. See ContentTypesByExtension for why this is an allowlist.
        //
        // The extension is read from the LAST dot, which is what the filesystem and every browser
        // agree on - reading the first would serve "evil.png.html" as an image and let the browser
        // sniff it back to HTML.
        // ###########################################################################################
        public static string ContentTypeFor(string? fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return ReviewAssetLocator.DefaultContentType;

            string extension = Path.GetExtension(fileName);

            if (string.IsNullOrEmpty(extension))
                return ReviewAssetLocator.DefaultContentType;

            return ReviewAssetLocator.ContentTypesByExtension.TryGetValue(extension, out string? contentType)
                ? contentType
                : ReviewAssetLocator.DefaultContentType;
        }
    }
}
