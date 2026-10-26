using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE CRT 4.0 DATA TREE - where everything lives in it (owner decisions, 2026-10-24: board data
    // as JSON only, every board owning its files, "having the files only saved once").
    //
    //   <tree>/                                  e.g. public_html/app-data-v4
    //     dataChecksums.json                     the manifest every CRT 4 syncs against
    //     Data/
    //       Classic-Repair-Toolbox.v4.0.0.json   the main file: the boards' list, the oscilloscopes
    //       <board id>/board.json                one board, everything about it (BoardDocument)
    //       Files/<ab>/<sha256>.<ext>            every file's bytes, stored ONCE
    //       MiniPro/IC tests/...                 CRT's own IC tests, read by folder
    //
    // *** A BOARD'S FOLDER IS WHERE IT ALWAYS WAS ("Commodore/C64/250407/"), holding only its
    // board.json. *** So a board's identity - the path CRT has always carried for it, its folder and
    // a file name - still finds its folder in the tree and its draft's folder under the drafts root;
    // only the file in it is now always "board.json" (BoardFileFor). "Files" and "MiniPro" are the
    // tree's own folders, so no manufacturer may be called either (the converter refuses one).
    //
    // *** A TREE OF ITS OWN. *** An installed 3.x downloads every file its manifest lists, so these
    // files never go beside the Excel generations (DataManager's sync) - the old trees stay as they
    // are for 2.x and 3.x, and this one is found at its own address.
    //
    // *** A BOARD'S FILES ARE ITS OWN, THE BYTES ARE SHARED. *** A board names each of its files
    // with a name of its own ("Component images/7406.png") and the file's SHA-256; the bytes are
    // stored once, under the checksum, however many boards name them. A changed file is new bytes
    // under a new checksum, so a stored file never changes and changing one board's file can never
    // change another board's.
    //
    // The main file keeps the generation in its name, as the Excel masters did: a later CRT that
    // needs another main file can add one beside this, and every CRT reads the newest at or below
    // its own version (DataGenerationRules).
    // ###########################################################################################
    public static class DataTreeLayout
    {
        // The folder beside the manifest that holds everything a CRT syncs - "Data", as in the
        // Excel trees, so a tree's public address ends ".../Data" in both.
        public const string ContentFolder = "Data";

        public const string ManifestFileName = "dataChecksums.json";

        public const string MainFileName = "Classic-Repair-Toolbox.v4.0.0.json";

        public const string BoardFileName = "board.json";

        public const string FilesFolder = "Files";

        public const string MiniProTestsFolder = "MiniPro/IC tests";

        // The tree's own top folders - never a manufacturer's.
        public static readonly IReadOnlyList<string> ReservedTopFolders = [DataTreeLayout.FilesFolder, "MiniPro"];

        // ###########################################################################################
        // The format of the board and main files. A reader refuses a file of a HIGHER version - a
        // later CRT changed something it cannot read - and keeps every field it does not know in a
        // file of its own version, so data a later CRT adds survives an edit made with this one.
        // ###########################################################################################
        public const int FormatVersion = 1;

        // "Commodore/C64/250407/board.json" - content-folder relative, "/" separated.
        public static string BoardFilePath(string boardId) =>
            $"{boardId.Trim().Trim('/')}/{DataTreeLayout.BoardFileName}";

        // ###########################################################################################
        // The board file of a board CRT identifies by `identity` - the path it has always carried
        // for a board ("Commodore/C64/250407/Data C64 250407.xlsx", or "Commodore/C64/250407/board.json"):
        // the same folder, "board.json" in it. Empty when the identity names no folder.
        // ###########################################################################################
        public static string BoardFileFor(string? identity)
        {
            string path = (identity ?? string.Empty).Trim().Replace('\\', '/').Trim('/');
            int slash = path.LastIndexOf('/');

            return slash <= 0 ? string.Empty : $"{path[..slash]}/{DataTreeLayout.BoardFileName}";
        }

        // Whether a path is a CRT 4.0 board file - the one test the shared readers choose the
        // format by (BoardDataReader, BoardComponentHighlightStorage): a "board.json" is read as
        // a BoardDocument, anything else as the Excel generations' workbook and sidecar.
        public static bool IsBoardFile(string? path) =>
            string.Equals(Path.GetFileName(path ?? string.Empty), DataTreeLayout.BoardFileName, StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Where the bytes of a file with this checksum are stored: "Files/3f/3f7a...e1.pdf".
        //
        // The two-character folder keeps any one folder small (a few thousand files spread over 256
        // folders). The extension is the board's name's, lower-cased, so the web server sends the
        // right type and a person looking at the folder can tell a picture from a datasheet; an
        // extension that is not plain letters, digits and "_" is left off rather than trusted.
        // Two boards naming the same bytes ".PDF" and ".pdf" share one stored file.
        // ###########################################################################################
        public static string StoredFilePath(string sha256, string name)
        {
            if (!DataTreeLayout.IsSha256(sha256))
                throw new ArgumentException($"[{sha256}] is not a lower-case SHA-256.", nameof(sha256));

            return $"{DataTreeLayout.FilesFolder}/{sha256[..2]}/{sha256}{DataTreeLayout.StoredExtension(name)}";
        }

        public static bool IsSha256(string? value) =>
            value is { Length: 64 } && value.All(c => c is (>= '0' and <= '9') or (>= 'a' and <= 'f'));

        // ###########################################################################################
        // The checksum a file's bytes are stored under - SHA-256 of the whole file, as lower-case hex.
        // ONE helper for the draft's store, the conversions and the tree (code review, 2026-10-25:
        // three copies of it had grown), because their checksums must compare equal.
        // ###########################################################################################
        public static string Sha256Of(string path)
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1 << 16);
            return Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(stream));
        }

        // ".pdf" for "Datasheets/7406.PDF", "" for a name with no usable extension.
        public static string StoredExtension(string? name)
        {
            string extension = Path.GetExtension(name ?? string.Empty).ToLowerInvariant();

            bool usable = extension.Length is >= 2 and <= 17 &&
                          extension.Skip(1).All(c => c is (>= 'a' and <= 'z') or (>= '0' and <= '9') or '_');

            return usable ? extension : string.Empty;
        }
    }
}
