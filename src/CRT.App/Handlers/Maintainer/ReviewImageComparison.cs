using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHICH images a maintainer is shown side by side (NewContributeStrategy.md Phase 5, task 4).
    //
    // The asset endpoint can fetch any of a submission's bytes; this decides which ones are worth
    // putting on screen and how each pair is described. Pure, so the decision that governs what a
    // maintainer actually looks at is unit tested rather than verified by opening the app.
    //
    // *** THE SUBMISSION'S FILE LIST IS THE COMPLETE INTENDED STATE, NOT A LIST OF CHANGES. ***
    // That is the manifest's own contract, and it is the single thing to understand here: an
    // untouched board re-lists every image it already had. Drawing all of them would bury the one
    // that moved, which is the same failure as opening on the whole board rather than a summary.
    // So a file is shown only when it is genuinely different from what is published.
    //
    // *** A DELETION IS FOUND FROM THE PUBLISHED SIDE, NEVER FROM THE MANIFEST. *** There is no
    // blob for a removed file, so anything driven off what was uploaded misses it completely. A
    // removal is the least recoverable thing a submission can do, so it is found by asking what
    // the published board HAS that this submission does not, and it is listed FIRST.
    // ###########################################################################################
    public static class ReviewImageComparison
    {

        // ###########################################################################################
        // The image comparisons worth showing, removals first.
        //
        // publishedHashesByPath is OPTIONAL. When it is supplied, a file whose hash matches what is
        // published is dropped as unchanged - which is most of them. When it is not, every path
        // present on both sides is treated as replaced, which is the safe direction to be wrong in:
        // showing a maintainer an unchanged pair wastes their time, whereas hiding a changed one
        // means it is approved unseen.
        //
        // Never throws. It is called while drawing a panel, from a submission whose payload may
        // not have loaded, and taking the review screen down over a missing list would be worse
        // than showing nothing.
        // ###########################################################################################
        public static IReadOnlyList<ReviewImagePair> Plan(
            ReviewSubmissionAssets? submitted,
            IReadOnlyCollection<string>? publishedImagePaths,
            IReadOnlyDictionary<string, string>? publishedHashesByPath = null)
        {
            // Ordinal throughout: the server's filesystem is Linux and the data tree is
            // case-sensitive from Phase 3 onward. Folding case here would pair an added file
            // against an unrelated published one and show a comparison between two different
            // images.
            var published = new HashSet<string>(
                publishedImagePaths?.Where(ReviewImageComparison.IsImage) ?? [],
                StringComparer.Ordinal);

            var pairs = new List<ReviewImagePair>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            // Removals first - see the header. Taken from the published side, because the
            // submission carries nothing at all for a file it is deleting.
            IEnumerable<string> submittedPaths = submitted?.Files
                .Select(file => file.Path)
                .Where(ReviewImageComparison.IsImage)
                ?? [];

            var submittedSet = new HashSet<string>(submittedPaths, StringComparer.Ordinal);

            foreach (string path in published.Where(path => !submittedSet.Contains(path)).OrderBy(path => path, StringComparer.Ordinal))
            {
                pairs.Add(new ReviewImagePair(
                    path,
                    ReviewImageChange.Removed,
                    SubmittedHash: string.Empty));
            }

            foreach (ReviewSubmittedFile file in submitted?.Files ?? [])
            {
                if (!ReviewImageComparison.IsImage(file.Path))
                    continue;

                // A manifest is contributor-supplied and can carry a duplicate row; two identical
                // panels would read as two different images that happen to look the same.
                if (!seen.Add(file.Path))
                    continue;

                // *** SOMETHING ALREADY PUBLISHED AT THIS PATH IS A REPLACEMENT, whether or not the
                // old board cited it (security review, 2026-09-25). *** The server now hashes every
                // submitted path that exists on disk. Judging by the old board's list alone called
                // a file that OVERWRITES an existing one "Added" - backwards for exactly the case a
                // maintainer most needs to see.
                bool existsBefore = published.Contains(file.Path) ||
                    (publishedHashesByPath?.ContainsKey(file.Path) ?? false);

                if (existsBefore
                    && publishedHashesByPath is not null
                    && publishedHashesByPath.TryGetValue(file.Path, out string? publishedHash)
                    && string.Equals(publishedHash, file.Sha256, StringComparison.Ordinal))
                {
                    // Byte-identical to what is already published. Not a change.
                    continue;
                }

                pairs.Add(new ReviewImagePair(
                    file.Path,
                    existsBefore ? ReviewImageChange.Replaced : ReviewImageChange.Added,
                    file.Sha256));
            }

            return pairs;
        }

        // ###########################################################################################
        // What the Maintainer tab will try to draw as a picture - CRT.Data's one list.
        //
        // *** A CONTRACT WITH ReviewAssetLocator's OWN ALLOWLIST on the server. *** Offering a
        // comparison for a type the server refuses to serve as an image produces a panel that
        // cannot decode what comes back. SVG is absent from both for the same reason: it is a
        // scriptable XML document, not a picture.
        // ###########################################################################################
        private static bool IsImage(string? path) => ImageFileTypes.IsDisplayable(path);
    }

    // ###########################################################################################
    // What a submission carries, as far as fetching bytes is concerned.
    //
    // A trimmed view of the manifest rather than the manifest itself: the window needs the path
    // and the hash and nothing else, and taking the whole contract type here would make every
    // future field on it look like something this screen might depend on.
    // ###########################################################################################
    public sealed record ReviewSubmissionAssets(IReadOnlyList<ReviewSubmittedFile> Files);

    public sealed record ReviewSubmittedFile(string Path, string Sha256);

    // ###########################################################################################
    // One image comparison, and which of the three things happened to it.
    //
    // The three are kept DISTINCT rather than collapsed into "changed", because they are
    // different decisions for the person looking - and because a missing "before" panel must be
    // captioned as "this is new" rather than left blank, which reads as a failure to load.
    // ###########################################################################################
    public sealed record ReviewImagePair(string Path, ReviewImageChange Change, string SubmittedHash)
    {
        public bool HasBefore => this.Change is ReviewImageChange.Replaced or ReviewImageChange.Removed;

        public bool HasAfter => this.Change is ReviewImageChange.Replaced or ReviewImageChange.Added;
    }

    public enum ReviewImageChange
    {
        Added,
        Replaced,
        Removed
    }
}
