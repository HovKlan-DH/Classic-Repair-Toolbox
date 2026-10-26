using System.Security.Cryptography;
using System.Text;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Writes a REAL CRT 4.0 data tree for the tests (DataTreeLayout): the main file listing boards, a
// board file per board listing its files by name and checksum (each cited by a board-file row, so
// the board's rows and its list agree), and each file's bytes stored once under their checksum.
// ###########################################################################################
internal static class JsonTree
{
    public static string Sha(string content) =>
        Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(content)));

    // "Files/ab/<sha>.<ext>" for the bytes of `content` named `name`.
    public static string Stored(string content, string name) =>
        DataTreeLayout.StoredFilePath(JsonTree.Sha(content), name);

    public static void MainFile(string root, params string[] boardIds)
    {
        var main = new MainDocument
        {
            Boards = boardIds.Select(id => new MainDocument.BoardListingRow
            {
                BoardId = id,
                HardwareName = id.Split('/')[1],
                BoardName = id.Split('/')[2]
            }).ToList()
        };

        JsonTree.Write(root, DataTreeLayout.MainFileName, DataTreeJson.Write(main));
    }

    // The board file of `boardId`, listing each file (name, content) and storing its bytes.
    public static string Board(string root, string boardId, params (string Name, string Content)[] files)
    {
        var document = new BoardDocument
        {
            RevisionDate = "2026-October-25",
            Components = [new BoardDocument.ComponentRow { BoardLabel = "U8", FriendlyName = "PLA" }],
            BoardFiles = files.Select(file => new BoardDocument.BoardFileRow { Category = "Service", Name = file.Name, File = file.Name }).ToList(),
            Files = files.Select(file => new BoardDocument.StoredFileRow
            {
                Name = file.Name,
                Sha256 = JsonTree.Sha(file.Content),
                Size = Encoding.UTF8.GetByteCount(file.Content)
            }).ToList()
        };

        foreach ((string name, string content) in files)
            JsonTree.Write(root, JsonTree.Stored(content, name), content);

        return JsonTree.Write(root, DataTreeLayout.BoardFilePath(boardId), DataTreeJson.Write(document));
    }

    public static string Write(string root, string relative, string content)
    {
        string full = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, content);
        return full;
    }
}
