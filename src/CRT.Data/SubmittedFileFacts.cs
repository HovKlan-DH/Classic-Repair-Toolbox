using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What the MAINTAINER needs to know about each file a submission carries (security review,
    // 2026-09-25).
    //
    // *** THE REVIEW SCREEN USED TO SHOW IMAGES ONLY. *** Everything else - a PDF, a text file, a
    // file no row used, a file belonging to another board - was drawn nowhere, so an
    // administrator approved it without ever being shown it. This is the per-file account the
    // server now sends alongside the change summary, and the Maintainer tab lists every file
    // that would change the tree.
    //
    // *** ONE TYPE FOR BOTH ENDS OF THE WIRE. *** The server serialises this record and the review
    // application deserialises the same record, so a renamed property cannot compile on one side
    // and silently read as blank on the other - the failure CLAUDE.md warns HTTP contracts hide.
    //
    // PublishedSha256 is the hash of whatever is published at this path NOW, not of the file the
    // published board references - a submission can name a path the old board never cited, and
    // overwriting an existing file there is exactly the case a maintainer must not see as "added".
    // ###########################################################################################
    public sealed record SubmittedFileFact(
        string Path,
        string Sha256,
        long SizeBytes,
        SubmissionFileScope Scope,
        bool IsReferenced,
        string? PublishedSha256)
    {
        public bool ExistsOnServer => this.PublishedSha256 is not null;

        public bool IsUnchanged =>
            this.PublishedSha256 is not null &&
            string.Equals(this.PublishedSha256, this.Sha256, StringComparison.Ordinal);
    }

    public static class SubmittedFileFacts
    {
        // ###########################################################################################
        // One fact per distinct submitted path, in path order.
        //
        // `publishedHashes` maps a relative path to the SHA-256 published there; a path absent from
        // it has nothing published at it.
        // ###########################################################################################
        public static IReadOnlyList<SubmittedFileFact> Build(
            SubmissionManifest manifest,
            IReadOnlyDictionary<string, string>? publishedHashes)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            IReadOnlySet<string> referenced = SubmissionFileRules.ReferencedFiles(manifest);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var facts = new List<SubmittedFileFact>();

            foreach (SubmissionFile file in manifest.Files.OrderBy(file => file.Path, StringComparer.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(file.Path) || !seen.Add(file.Path))
                    continue;

                string? published = null;
                publishedHashes?.TryGetValue(file.Path, out published);

                facts.Add(new SubmittedFileFact(
                    file.Path,
                    file.Sha256,
                    file.SizeBytes,
                    SubmissionFileScopes.Classify(manifest, file.Path),
                    referenced.Contains(file.Path),
                    published));
            }

            return facts;
        }
    }
}
