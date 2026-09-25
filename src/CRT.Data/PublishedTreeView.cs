using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What the PUBLISHED tree holds, as far as the submission rules need to know (security review,
    // 2026-09-25).
    //
    // Two questions, both about RELATIVE paths ("Commodore/C64/250407/Sheet1.png"):
    //
    //   HashOf         - the SHA-256 of the file published at this path, or null when there is
    //                    none. Lets a submission cite another board's file UNCHANGED (see
    //                    SubmissionFileScope) and lets the reviewer be told a file already exists.
    //   ListDirectory  - the names of the entries in a folder ("" is the data root), or null when
    //                    the folder does not exist. Lets a case-only variant be caught.
    //
    // *** DELEGATES, SO THE RULES STAY PURE. *** The server answers these from its data tree;
    // tests answer them from a dictionary. The rules in SubmissionFileRules and PublishPlan never
    // touch a disk, so every one of them is a unit test.
    //
    // *** WHY CASE VARIANTS MATTER ENOUGH TO LOOK FOR. *** The server's filesystem is Linux, where
    // "commodore/C64/250407" and "Commodore/C64/250407" are two folders. Every Windows and macOS
    // client folds them into ONE, so whichever manifest entry syncs last overwrites the other - a
    // submission to a case-variant "new system" would replace a real board's images on every
    // client without the real board's files on the server ever being touched.
    // ###########################################################################################
    public sealed class PublishedTreeView
    {
        private readonly Func<string, string?> thisHashOf;
        private readonly Func<string, IReadOnlyCollection<string>?> thisListDirectory;

        // One listing per folder per view. A 1,200-file submission walks the same few folders
        // over and over, and each listing is a real directory read on the server.
        private readonly Dictionary<string, IReadOnlyCollection<string>?> thisListings =
            new(StringComparer.Ordinal);

        public PublishedTreeView(
            Func<string, string?> hashOf,
            Func<string, IReadOnlyCollection<string>?> listDirectory)
        {
            ArgumentNullException.ThrowIfNull(hashOf);
            ArgumentNullException.ThrowIfNull(listDirectory);

            this.thisHashOf = hashOf;
            this.thisListDirectory = listDirectory;
        }

        // A view of an EMPTY tree - nothing published, nothing to collide with. What a brand-new
        // data root looks like, and a convenient default for tests.
        public static PublishedTreeView Empty { get; } = new(_ => null, _ => null);

        // ###########################################################################################
        // The lowercase SHA-256 of the file published at this relative path, or null.
        // ###########################################################################################
        public string? HashOf(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                return null;

            return this.thisHashOf(relativePath);
        }

        // ###########################################################################################
        // The EXISTING path that differs from this one only by capitalisation, or null when there
        // is none.
        //
        // Walks the path one segment at a time from the root. At each level: the exact name
        // existing means keep walking; a case-only match means that is the collision; neither means
        // the path is new from here down and nothing deeper can collide.
        //
        // Returns the colliding path AS IT EXISTS, so the message can name both spellings.
        // ###########################################################################################
        public string? FindCaseVariant(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                return null;

            string[] segments = relativePath.Split('/');
            string prefix = string.Empty;

            foreach (string segment in segments)
            {
                IReadOnlyCollection<string>? entries = this.List(prefix);

                if (entries is null)
                    return null;

                string current = prefix.Length == 0 ? segment : prefix + "/" + segment;

                if (entries.Contains(segment, StringComparer.Ordinal))
                {
                    prefix = current;
                    continue;
                }

                string? variant = entries.FirstOrDefault(
                    entry => string.Equals(entry, segment, StringComparison.OrdinalIgnoreCase));

                if (variant is not null)
                    return prefix.Length == 0 ? variant : prefix + "/" + variant;

                return null;
            }

            return null;
        }

        private IReadOnlyCollection<string>? List(string relativeFolder)
        {
            if (!this.thisListings.TryGetValue(relativeFolder, out IReadOnlyCollection<string>? entries))
            {
                entries = this.thisListDirectory(relativeFolder);
                this.thisListings[relativeFolder] = entries;
            }

            return entries;
        }
    }
}
