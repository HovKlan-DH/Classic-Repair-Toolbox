using System;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Does a submission ADD or CHANGE a shared file? (Phase 6 roles, 2026-09-25.)
    //
    // *** SHARED FILES ARE ADMINISTRATOR-OWNED. *** "Commodore/Shared files" and "Generic shared
    // files" are cited by many boards, so a change there reaches every one of them - which is why
    // NewContributeStrategy.md's Phase 6 traps say a submission touching them "must not be
    // approvable by a system maintainer". A reviewer is assigned to a SYSTEM; a shared file belongs
    // to none.
    //
    // A shared file cited UNCHANGED does not count: that is the ordinary shape of a board that uses
    // a shared component image, and refusing it to reviewers would send nearly every submission to
    // the administrator. Only bytes that would be written matter - a path with nothing published at
    // it, or a hash that differs from what is.
    //
    // *** NO VIEW OF THE TREE MEANS "CHANGED". *** When the published tree could not be consulted
    // the hash cannot be compared, and "administrator only" is the safe side to fall on - the same
    // direction SubmissionFileRules takes for a foreign file it cannot check.
    //
    // Pure, and in CRT.Data so the same rule can be run by any side that holds a manifest.
    // ###########################################################################################
    public static class SubmissionSharedFiles
    {
        public static bool TouchesSharedFiles(SubmissionManifest manifest, PublishedTreeView? tree)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return manifest.Files.Any(file => SubmissionSharedFiles.IsChangedSharedFile(manifest, file, tree));
        }

        private static bool IsChangedSharedFile(SubmissionManifest manifest, SubmissionFile file, PublishedTreeView? tree)
        {
            SubmissionFileScope scope = SubmissionFileScopes.Classify(manifest, file.Path);

            if (scope is not (SubmissionFileScope.ManufacturerShared or SubmissionFileScope.GenericShared))
                return false;

            string? published = tree?.HashOf(file.Path);

            return published is null || !string.Equals(published, file.Sha256, StringComparison.Ordinal);
        }
    }
}
