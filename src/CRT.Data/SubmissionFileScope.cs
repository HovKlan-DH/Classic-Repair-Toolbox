using System;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHOSE file a submitted path is (security review, 2026-09-25).
    //
    // *** THIS IS WHAT STOPS A SUBMISSION TO ONE BOARD REWRITING ANOTHER. *** Submitted paths are
    // contained to the DATA ROOT, not to the submission's own folder, because a board legitimately
    // cites shared files that sit beside the manufacturer ("Commodore/Shared files/...") or at the
    // top of the tree ("Generic shared files/..."). Containment alone therefore let a submission to
    // the Amstrad CPC carry "Commodore/C64/250407/<anything>" - and the publish copied it over the
    // real file, with nothing on the review screen to show it had happened.
    //
    // A submission may CHANGE only three places:
    //
    //   Own                - "<Manufacturer>/<Hardware>/<Board>/..." - the system it names.
    //   ManufacturerShared - "<Manufacturer>/Shared files/..." - shared by that maker's boards.
    //   GenericShared      - "Generic shared files/..." - shared by every board.
    //
    // Anything else is Foreign. A foreign path is not automatically wrong: the published C128DCR
    // 250477 board cites two scope-baseline texts that live in the C128 310378 folder, so a strict
    // "own folder only" rule rejects correct, already-published data (found by reading every
    // shipped board, not guessed). What a submission may NOT do is CHANGE a foreign file - see
    // SubmissionFileRules, which allows a foreign path only when its bytes equal what is already
    // published there.
    //
    // CASE-SENSITIVE, like every path comparison against the Linux tree: "commodore/C64/250407/"
    // is not this system's folder, and treating it as one would let a case-variant pass as own.
    //
    // Serialised as its NAME, not its number, because the review application reads it off the
    // wire - a number that shifted when a member was added would mislabel every file silently.
    // ###########################################################################################
    [JsonConverter(typeof(JsonStringEnumConverter<SubmissionFileScope>))]
    public enum SubmissionFileScope
    {
        Own,
        ManufacturerShared,
        GenericShared,
        Foreign
    }

    public static class SubmissionFileScopes
    {
        // The folder beside a manufacturer's boards that all of them may cite.
        public const string SharedFilesFolderName = "Shared files";

        // The top-level folder every board may cite.
        public const string GenericSharedFolderName = "Generic shared files";

        // ###########################################################################################
        // Is this folder name one of the two SHARED folders? The ONE rule for it (code review,
        // 2026-09-25): SubmissionValidator refuses a board named with either name in either of the
        // first two positions, and DataTreeUsage and PublishedSystemLister skip exactly those
        // folders. Three hand-written copies had drifted apart - the validator checked one name per
        // position - so a board publishable under "Shared files/..." was invisible to the other two.
        //
        // Case-insensitive: every Windows and macOS client folds "shared FILES" into "Shared files".
        // ###########################################################################################
        public static bool IsSharedFolderName(string? name) =>
            string.Equals(name, SubmissionFileScopes.SharedFilesFolderName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, SubmissionFileScopes.GenericSharedFolderName, StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Which of the four a path is, for a submission naming this manufacturer, hardware and
        // board.
        //
        // A blank identity part classifies everything as Foreign rather than matching a prefix of
        // "//" - an unidentifiable submission owns nothing. The validator refuses such a manifest
        // anyway; this only guarantees the classification cannot be the thing that lets it through.
        // ###########################################################################################
        public static SubmissionFileScope Classify(
            string? manufacturer,
            string? hardware,
            string? board,
            string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return SubmissionFileScope.Foreign;

            bool hasManufacturer = !string.IsNullOrWhiteSpace(manufacturer);

            if (hasManufacturer &&
                !string.IsNullOrWhiteSpace(hardware) &&
                !string.IsNullOrWhiteSpace(board) &&
                path.StartsWith($"{manufacturer}/{hardware}/{board}/", StringComparison.Ordinal))
            {
                return SubmissionFileScope.Own;
            }

            if (hasManufacturer &&
                path.StartsWith(
                    $"{manufacturer}/{SubmissionFileScopes.SharedFilesFolderName}/",
                    StringComparison.Ordinal))
            {
                return SubmissionFileScope.ManufacturerShared;
            }

            if (path.StartsWith(SubmissionFileScopes.GenericSharedFolderName + "/", StringComparison.Ordinal))
                return SubmissionFileScope.GenericShared;

            return SubmissionFileScope.Foreign;
        }

        // The same, for a manifest.
        public static SubmissionFileScope Classify(SubmissionManifest manifest, string? path)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return SubmissionFileScopes.Classify(manifest.Manufacturer, manifest.Hardware, manifest.Board, path);
        }
    }
}
