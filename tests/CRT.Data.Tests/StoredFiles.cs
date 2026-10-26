using System.Security.Cryptography;
using System.Text;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Builds the pieces of a CRT 4.0 tree a test needs: a file's bytes in a store
// (DataTreeLayout.StoredFilePath under a root - the tree's or a draft folder's) and a board file
// listing files by name.
// ###########################################################################################
internal static class StoredFiles
{
    // Stores `content` under `root` as the file `name` and returns its entry.
    public static BoardDocument.StoredFileRow Put(string root, string name, string content)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        string sha = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        string path = Path.Combine(root, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, bytes);

        return new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.Length };
    }

    // An entry for `content` without storing it anywhere.
    public static BoardDocument.StoredFileRow EntryFor(string name, string content)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        return new BoardDocument.StoredFileRow
        {
            Name = name,
            Sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
            Size = bytes.Length,
        };
    }

    // Writes a board file listing `files` and returns its path.
    public static string WriteBoard(string boardFile, params BoardDocument.StoredFileRow[] files)
    {
        BoardFiles.Write(boardFile, new BoardDocument { Files = files.ToList() });
        return boardFile;
    }
}
