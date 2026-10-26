using System.Security.Cryptography;
using System.Text;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The pieces of a CRT 4.0 tree a test needs: a file's bytes in a store (DataTreeLayout
// .StoredFilePath under a root - the tree's or a draft folder's) and a board file naming its files.
// The same helper CRT.Data.Tests has.
// ###########################################################################################
internal static class StoredFiles
{
    // Stores `content` under `root` as the file `name` and returns its entry.
    public static BoardDocument.StoredFileRow Put(string root, string name, string content) =>
        StoredFiles.Put(root, name, Encoding.UTF8.GetBytes(content));

    // The same, for bytes that are not text - a picture a test has CRT draw.
    public static BoardDocument.StoredFileRow Put(string root, string name, byte[] bytes)
    {
        string sha = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

        Directory.CreateDirectory(Path.GetDirectoryName(StoredFiles.PathOf(root, sha, name))!);
        File.WriteAllBytes(StoredFiles.PathOf(root, sha, name), bytes);

        return new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.Length };
    }

    // Where `file`'s bytes are stored under `root`.
    public static string PathOf(string root, BoardDocument.StoredFileRow file) =>
        StoredFiles.PathOf(root, file.Sha256, file.Name);

    private static string PathOf(string root, string sha, string name) =>
        Path.Combine(root, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));

    // Writes `board` as the board file at `boardFile`, naming `files`.
    public static string WriteBoard(string boardFile, BoardData board, params BoardDocument.StoredFileRow[] files)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(boardFile))!);
        BoardWorkbookWriter.Write(boardFile, board);
        BoardFiles.Change(boardFile, document => document.Files = files.ToList());
        return boardFile;
    }
}
