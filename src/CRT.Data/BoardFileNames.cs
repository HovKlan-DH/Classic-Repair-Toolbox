using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE NAMES A BOARD GIVES ITS FILES IN THE CRT 4.0 DATA TREE (owner decision, 2026-10-24: no
    // generic or shared files - every board owns all its files).
    //
    // A row in the Excel generations cites a file by its path in the tree. In 4.0 a board names
    // each of its files itself, and the name is that path without the part that said WHERE it was
    // kept:
    //
    //   the board's own folder      Commodore/C64/250407/Scope baseline/U19_1_PAL.png  -> Scope baseline/U19_1_PAL.png
    //   a manufacturer's shared     Commodore/Shared files/Component images/x.png      -> Component images/x.png
    //   the generic shared files    Generic shared files/Component local files/7406.PDF -> Component local files/7406.PDF
    //   another board's folder      Commodore/C64/250425/Board.png                     -> Board.png
    //   anything else               kept as it is
    //
    // *** TWO FILES, ONE NAME. *** Names are compared ignoring capitals, as Windows and macOS do.
    // When a second path would get a name already given to OTHER bytes, it gets the next free
    // numbered name ("x (2).png"); a second path to the SAME bytes shares the name - it is the same
    // file. The first path to ask keeps the plain name, so the converter asks in the order the
    // board shows its files.
    //
    // *** A PATH IS ITS FILE WHATEVER ITS SPELLING. *** "Board\U8.png", "/Board/U8.png" and
    // "Board/./x/../U8.png" all name "Board/U8.png", as every client's file system resolves them -
    // DataTreeUsage counts them so too - and a path is matched ignoring capitals.
    // ###########################################################################################
    public sealed class BoardFileNames
    {
        private const string GenericSharedFolder = "Generic shared files";
        private const string ManufacturerSharedFolder = "Shared files";

        private readonly string thisBoardId;
        private readonly IReadOnlyList<string> thisOtherBoardIds;
        private readonly Dictionary<string, string> thisNameByPath = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> thisShaByName = new(StringComparer.OrdinalIgnoreCase);

        // `boardIds` is every board the tree holds, this one included - a file cited from another
        // board's folder loses that board's folders too. Longest first, so a board whose id begins
        // with another's is matched as itself.
        public BoardFileNames(string boardId, IEnumerable<string> boardIds)
        {
            this.thisBoardId = BoardFileNames.Normalise(boardId);
            this.thisOtherBoardIds = (boardIds ?? [])
                .Select(BoardFileNames.Normalise)
                .Where(id => id.Length > 0 && !string.Equals(id, this.thisBoardId, StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(id => id.Length)
                .ThenBy(id => id, StringComparer.Ordinal)
                .ToList();
        }

        // Every path named so far, and the name it got.
        public IReadOnlyDictionary<string, string> NamesByPath => this.thisNameByPath;

        // ###########################################################################################
        // The name of the file at `citedPath` (data-root relative, as a row cites it), whose bytes
        // have checksum `sha256` - given now, or the one it was given before.
        // ###########################################################################################
        public string Add(string citedPath, string sha256)
        {
            string path = BoardFileNames.Normalise(citedPath);

            if (path.Length == 0)
                throw new ArgumentException("A file needs a path.", nameof(citedPath));

            if (this.thisNameByPath.TryGetValue(path, out string? known))
                return known;

            string preferred = this.PreferredName(path);
            string name = preferred;

            for (int number = 2; this.thisShaByName.TryGetValue(name, out string? taken) && !string.Equals(taken, sha256, StringComparison.Ordinal); number++)
                name = BoardFileNames.Numbered(preferred, number);

            this.thisShaByName.TryAdd(name, sha256);
            this.thisNameByPath[path] = name;

            return name;
        }

        // ###########################################################################################
        // A name another board file already has, for these bytes - the PUBLISHED board's names
        // when a draft is converted beside it (DraftConversion). A new file asking for the name is
        // numbered, as for any other taken name; it is not given to any path by this.
        // ###########################################################################################
        public void Reserve(string name, string sha256) => this.thisShaByName.TryAdd(name, sha256);

        // ###########################################################################################
        // Gives the path this name outright, its bytes now being `sha256` - a file the draft
        // REPLACED keeps the name the published board knows it by, so it reads as replaced rather
        // than as one file removed and another added.
        // ###########################################################################################
        public void Assign(string citedPath, string name, string sha256)
        {
            string path = BoardFileNames.Normalise(citedPath);

            if (path.Length == 0 || string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("A file needs a path and a name.", nameof(citedPath));

            this.thisNameByPath[path] = name;
            this.thisShaByName[name] = sha256;
        }

        // The name a path was given, or the one it would get - for a reference to a file the board
        // no longer holds (a rename's image that was replaced since). Registers nothing.
        public string NameFor(string? citedPath)
        {
            string path = BoardFileNames.Normalise(citedPath);

            if (path.Length == 0)
                return string.Empty;

            return this.thisNameByPath.TryGetValue(path, out string? known) ? known : this.PreferredName(path);
        }

        // The rule in the class header, without the numbering.
        public string PreferredName(string normalisedPath)
        {
            string path = BoardFileNames.Normalise(normalisedPath);

            if (BoardFileNames.TryStripFolder(path, this.thisBoardId, out string rest))
                return rest;

            if (BoardFileNames.TryStripFolder(path, BoardFileNames.GenericSharedFolder, out rest))
                return rest;

            string[] segments = path.Split('/');

            if (segments.Length > 2 && string.Equals(segments[1], BoardFileNames.ManufacturerSharedFolder, StringComparison.OrdinalIgnoreCase))
                return string.Join('/', segments.Skip(2));

            foreach (string other in this.thisOtherBoardIds)
            {
                if (BoardFileNames.TryStripFolder(path, other, out rest))
                    return rest;
            }

            return path;
        }

        // ###########################################################################################
        // A path as a client's file system resolves it: trimmed, "/" separators, no leading slash,
        // "." segments and empty segments dropped, "x/.." folded away. A ".." that would climb out
        // of the tree is kept, so the file is not found rather than found somewhere else.
        // ###########################################################################################
        public static string Normalise(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            var segments = new List<string>();

            foreach (string segment in path.Trim().Replace('\\', '/').Split('/'))
            {
                if (segment.Length == 0 || segment == ".")
                    continue;

                if (segment == ".." && segments.Count > 0 && segments[^1] != "..")
                {
                    segments.RemoveAt(segments.Count - 1);
                    continue;
                }

                segments.Add(segment);
            }

            return string.Join('/', segments);
        }

        private static bool TryStripFolder(string path, string folder, out string rest)
        {
            rest = string.Empty;

            if (folder.Length == 0 ||
                path.Length <= folder.Length + 1 ||
                path[folder.Length] != '/' ||
                !path.StartsWith(folder, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            rest = path[(folder.Length + 1)..];
            return true;
        }

        // "Component images/x (2).png" for "Component images/x.png" and 2.
        private static string Numbered(string name, int number)
        {
            int slash = name.LastIndexOf('/');
            string folder = slash < 0 ? string.Empty : name[..(slash + 1)];
            string file = slash < 0 ? name : name[(slash + 1)..];
            int dot = file.LastIndexOf('.');

            return dot > 0
                ? $"{folder}{file[..dot]} ({number}){file[dot..]}"
                : $"{folder}{file} ({number})";
        }
    }
}
