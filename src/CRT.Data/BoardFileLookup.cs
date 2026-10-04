using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHETHER A FILE A ROW NAMES IS THERE - asked by BoardDataChecks, answered differently by
    // whoever runs the checks:
    //
    //   - the SERVER (SuppliedFileLookup): the paths the submission carries, compared exactly as
    //     the server's filesystem will, with a case-insensitive second look only to say "spelled
    //     differently" rather than "missing" - SubmissionValidator's rule since Phase 3;
    //   - the DRAFTS TAB and the launch log (DiskFileLookup): the files on this computer, found the
    //     way a submit finds them (SubmissionFileLocator - the draft's own copy first, then the
    //     downloaded data).
    //
    // `Where` is the phrase the message uses for that place, so the server's words stay exactly as
    // they were ("which is not in the submission") and the table's read naturally ("which is not
    // on this computer").
    // ###########################################################################################
    public interface IBoardFileLookup
    {
        string Where { get; }

        BoardFileLookupResult Check(string path);
    }

    public enum BoardFileState
    {
        Found,

        // Not there at all.
        Missing,

        // There, but spelled with different capitalisation - Actual says how.
        CaseDiffers
    }

    public readonly record struct BoardFileLookupResult(BoardFileState State, string Actual = "")
    {
        public static BoardFileLookupResult Found { get; } = new(BoardFileState.Found);

        public static BoardFileLookupResult Missing { get; } = new(BoardFileState.Missing);
    }

    // ###########################################################################################
    // The server's view: the set of paths a submission supplies. Ordinal first, as the server's
    // filesystem compares; the case-insensitive copy is ONLY for the better message - see
    // SubmissionValidator's header on why making the comparison itself case-insensitive would
    // accept genuinely broken data.
    // ###########################################################################################
    public sealed class SuppliedFileLookup : IBoardFileLookup
    {
        private readonly HashSet<string> thisExact;
        private readonly Dictionary<string, string> thisIgnoringCase = new(StringComparer.OrdinalIgnoreCase);

        public SuppliedFileLookup(IEnumerable<string> suppliedPaths)
        {
            ArgumentNullException.ThrowIfNull(suppliedPaths);

            List<string> paths = suppliedPaths.ToList();
            this.thisExact = new HashSet<string>(paths, StringComparer.Ordinal);

            foreach (string path in paths)
                this.thisIgnoringCase.TryAdd(path, path);
        }

        public string Where => "in the submission";

        public BoardFileLookupResult Check(string path)
        {
            if (this.thisExact.Contains(path))
                return BoardFileLookupResult.Found;

            return this.thisIgnoringCase.TryGetValue(path, out string? actual)
                ? new BoardFileLookupResult(BoardFileState.CaseDiffers, actual)
                : BoardFileLookupResult.Missing;
        }
    }

    // ###########################################################################################
    // This computer's view: a draft's own folder, then the downloaded data - SubmissionFileLocator,
    // the very rule a submit uses to find the bytes it uploads, so "found" here means the submit
    // finds it too.
    //
    // *** A FILE IN THE DOWNLOADED DATA MUST BE SPELLED EXACTLY. *** Windows and macOS find
    // "u8.png" when the file is "U8.png"; the server, on Linux, does not - and it refuses a path
    // differing only by capitalisation from a published one (SubmissionFileRules'
    // path.case_collision_published). So a match there is checked segment by segment against the
    // real names, which also finds it on a case-sensitive disk where File.Exists alone says
    // "missing". A file in the DRAFT folder is uploaded under the name the row gives it, so its
    // spelling on disk does not matter.
    //
    // A FOUND answer is kept for the life of the lookup - one table, one launch check - since the
    // table asks again after every edit. Folder listings are kept too: a board's files sit in a
    // handful of folders.
    //
    // *** A FILE THAT WAS NOT THERE IS LOOKED FOR AGAIN (code review, 2026-10-04). *** Every answer
    // was kept, so a "file missing" error stayed on the open table's cell after the picture was put
    // in the draft folder - while the draft's row (DraftStatusReader.CountProblemsCached, which
    // already asks such paths again) and Submit no longer found anything wrong. A path not found as
    // cited is resolved afresh on every ask, and a folder listing it reads is read again when the
    // folder has been written since. Usually there are none, so it costs nothing. A file that
    // VANISHES is still only seen by a new lookup - the cached count's rule too: re-asking every
    // found file (thousands on a large board) after every edit is the cost this cache avoids.
    // ###########################################################################################
    public sealed class DiskFileLookup : IBoardFileLookup
    {
        private readonly string thisDataRoot;
        private readonly string thisDraftSystemFolder;
        private readonly HashSet<string> thisFound = new(StringComparer.Ordinal);
        private readonly HashSet<string> thisAskedAndMissed = new(StringComparer.Ordinal);
        private readonly Dictionary<string, (DateTime Written, string[] Names)> thisListings = new(StringComparer.Ordinal);

        // `draftSystemFolder` is empty for the downloaded data alone (the launch log).
        public DiskFileLookup(string dataRoot, string draftSystemFolder)
        {
            this.thisDataRoot = dataRoot ?? string.Empty;
            this.thisDraftSystemFolder = draftSystemFolder ?? string.Empty;
        }

        public string Where => "on this computer";

        public BoardFileLookupResult Check(string path)
        {
            if (this.thisFound.Contains(path))
                return BoardFileLookupResult.Found;

            // Asked before and not found: the folders it looks in may have been written since.
            BoardFileLookupResult answer = this.Resolve(path, again: this.thisAskedAndMissed.Contains(path));

            if (answer.State == BoardFileState.Found)
            {
                this.thisFound.Add(path);
                this.thisAskedAndMissed.Remove(path);
            }
            else
            {
                this.thisAskedAndMissed.Add(path);
            }

            return answer;
        }

        private BoardFileLookupResult Resolve(string path, bool again)
        {
            bool located = SubmissionFileLocator.TryLocate(
                this.thisDataRoot, this.thisDraftSystemFolder, path, out string absolute, out _);

            if (located && !this.IsUnderDataRoot(absolute))
                return BoardFileLookupResult.Found;

            if (this.thisDataRoot.Length == 0)
                return located ? BoardFileLookupResult.Found : BoardFileLookupResult.Missing;

            // In the downloaded data - or nowhere, which on a case-sensitive disk can still be a
            // file spelled differently there.
            string? actual = this.ActualSpellingInDataRoot(path, again);

            if (actual is null)
                return located ? BoardFileLookupResult.Found : BoardFileLookupResult.Missing;

            return string.Equals(actual, path, StringComparison.Ordinal)
                ? BoardFileLookupResult.Found
                : new BoardFileLookupResult(BoardFileState.CaseDiffers, actual);
        }

        private bool IsUnderDataRoot(string absolute)
        {
            if (this.thisDataRoot.Length == 0)
                return false;

            string root = Path.GetFullPath(this.thisDataRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;

            return Path.GetFullPath(absolute).StartsWith(root, StringComparison.OrdinalIgnoreCase);
        }

        // The path as the data root's real names spell it, matching each segment ignoring case -
        // null when some segment is not there at all. `again`: a path asked about before, whose
        // folders' listings are checked against their write time rather than trusted.
        private string? ActualSpellingInDataRoot(string path, bool again)
        {
            string current = this.thisDataRoot;
            var actual = new List<string>();

            foreach (string segment in path.Split('/', StringSplitOptions.RemoveEmptyEntries))
            {
                string[] listing = this.ListingOf(current, again);

                string? match = listing.FirstOrDefault(name => string.Equals(name, segment, StringComparison.Ordinal))
                    ?? listing.FirstOrDefault(name => string.Equals(name, segment, StringComparison.OrdinalIgnoreCase));

                if (match is null)
                    return null;

                actual.Add(match);
                current = Path.Combine(current, match);
            }

            return actual.Count == 0 ? null : string.Join('/', actual);
        }

        private string[] ListingOf(string folder, bool checkWritten)
        {
            bool known = this.thisListings.TryGetValue(folder, out (DateTime Written, string[] Names) cached);

            if (known && !checkWritten)
                return cached.Names;

            DateTime written = DiskFileLookup.WrittenAt(folder);

            if (known && cached.Written == written)
                return cached.Names;

            string[] names;

            try
            {
                names = Directory.Exists(folder)
                    ? Directory.EnumerateFileSystemEntries(folder).Select(entry => Path.GetFileName(entry)).ToArray()
                    : [];
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                names = [];
            }

            this.thisListings[folder] = (written, names);
            return names;
        }

        // When a folder's entries last changed - a file added or removed writes its folder. MinValue
        // for one that is not there or cannot be read.
        private static DateTime WrittenAt(string folder)
        {
            try
            {
                return Directory.Exists(folder) ? Directory.GetLastWriteTimeUtc(folder) : DateTime.MinValue;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return DateTime.MinValue;
            }
        }
    }
}
