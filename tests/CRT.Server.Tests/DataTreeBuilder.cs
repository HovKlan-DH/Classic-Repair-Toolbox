using System.Security.Cryptography;
using System.Text;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Writes a REAL CRT 4.0 data tree for the server's tests (DataTreeLayout, 2026-10-25): the main
    // file listing boards, a board file per board listing its files by name and checksum - each
    // cited by a board-file row, so the board's rows and its list agree - and each file's bytes
    // stored once under their checksum. DataTreeUsage, the publish and the promotion read every one
    // of these, so a placeholder would make the whole tree unreadable - and an unreadable tree
    // removes nothing, which is the one thing these tests need to see NOT happen.
    // ###########################################################################################
    internal static class DataTreeBuilder
    {
        public const string BoardId = "Commodore/C64/250407";

        // "Commodore/C64/250407/board.json" - the board file of BoardId.
        public const string BoardFile = "Commodore/C64/250407/board.json";

        public static string Sha(string content) => DataTreeBuilder.Sha(Encoding.UTF8.GetBytes(content));

        public static string Sha(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

        // "Files/ab/<sha>.<ext>" for these bytes under this name.
        public static string Stored(string content, string name) => DataTreeLayout.StoredFilePath(DataTreeBuilder.Sha(content), name);

        // The main file, listing the given boards (ids, or board files) in order.
        public static void Master(string root, params string[] listed) =>
            DataTreeBuilder.ListingMaster(
                root,
                listed.Select(entry => BoardDescriptorRules.BoardIdFromExcelDataFile(entry) is { Length: > 0 } id ? id : entry.Trim('/'))
                    .Select(id => new MasterListingRow(id.Split('/')[1], id.Split('/')[2], DataTreeLayout.BoardFilePath(id), string.Empty))
                    .ToArray());

        // ###########################################################################################
        // The main file listing these rows in order - what a NEW board's publish adds its row to
        // (2026-09-27). The oscilloscopes are left empty.
        // ###########################################################################################
        public static void ListingMaster(string root, params MasterListingRow[] rows)
        {
            var main = new MainDocument
            {
                Boards = rows.Select(row => new MainDocument.BoardListingRow
                {
                    BoardId = row.BoardId,
                    HardwareName = row.HardwareName,
                    BoardName = row.BoardName,
                    HardwareNotes = row.Notes
                }).ToList()
            };

            File.WriteAllText(DataTreeBuilder.Full(root, DataTreeLayout.MainFileName), DataTreeJson.Write(main));
        }

        // The rows of a tree's main file, top to bottom.
        public static IReadOnlyList<MasterListingRow> ListedIn(string root)
        {
            Assert.True(MasterListing.TryRead(DataTreeBuilder.Full(root, DataTreeLayout.MainFileName), out IReadOnlyList<MasterListingRow> rows, out string why), why);
            return rows;
        }

        // ###########################################################################################
        // A board file whose "Board local files" rows cite the given files (name, content), with one
        // component - each file's bytes stored under their checksum. Returns the board file's path.
        // ###########################################################################################
        public static string Board(string root, string boardId, params (string Name, string Content)[] files) =>
            DataTreeBuilder.BoardWithBytes(root, boardId, files.Select(file => (file.Name, Encoding.UTF8.GetBytes(file.Content))).ToArray());

        public static string BoardWithBytes(string root, string boardId, params (string Name, byte[] Bytes)[] files)
        {
            var data = new BoardData
            {
                RevisionDate = "2026-August-21",
                Components =
                [
                    new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-01" }
                ]
            };

            foreach ((string name, _) in files)
                data.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = Path.GetFileName(name), File = name });

            return DataTreeBuilder.Board(root, boardId, data, files);
        }

        // A board file holding `board`, listing `files` and storing their bytes.
        public static string Board(string root, string boardId, BoardData board, IEnumerable<(string Name, byte[] Bytes)> files)
        {
            var listed = new List<BoardDocument.StoredFileRow>();

            foreach ((string name, byte[] bytes) in files)
            {
                string sha = DataTreeBuilder.Sha(bytes);
                File.WriteAllBytes(DataTreeBuilder.Full(root, DataTreeLayout.StoredFilePath(sha, name)), bytes);
                listed.Add(new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.LongLength });
            }

            BoardDocument document = BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, listed);
            string path = DataTreeBuilder.Full(root, DataTreeLayout.BoardFilePath(boardId));

            BoardFiles.Write(path, document);

            return path;
        }

        // Bytes stored under their checksum that no board lists.
        public static string StoreOnly(string root, string name, string content)
        {
            string relative = DataTreeBuilder.Stored(content, name);
            File.WriteAllText(DataTreeBuilder.Full(root, relative), content);
            return relative;
        }

        public static string Full(string root, string relative)
        {
            string full = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            return full;
        }
    }
}
