using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A CRT 4.0 data root with ONE published board and its pictures - the main file listing it, its
// board file, and each picture's bytes in the tree's store under their checksum - loaded into
// DataManager. For a surface that must find a board's pictures as CRT 4.0 lays them out, where a
// picture row NAMES a file of its board ("Images/top.png") instead of a path under the data root.
//
// It sets DataManager's static data root: a test using it puts that back with Unload.
// ###########################################################################################
internal static class PublishedBoardTree
{
    // Loads `dataRoot` listing the board `boardKey` ("Hardware|Board") as `board`, with `pictures`
    // (name -> bytes) stored; returns where each picture's bytes lie, by name.
    public static Dictionary<string, string> Load(string dataRoot, string boardKey, BoardData board, params (string Name, byte[] Bytes)[] pictures)
    {
        int bar = boardKey.IndexOf('|', StringComparison.Ordinal);
        var listed = new MainFiles.Board("Test/" + Guid.NewGuid().ToString("N"), boardKey[..bar], boardKey[(bar + 1)..]);

        DataManager.LoadFrom(dataRoot, MainFiles.Write(dataRoot, listed));

        BoardDocument.StoredFileRow[] rows = pictures.Select(picture => StoredFiles.Put(dataRoot, picture.Name, picture.Bytes)).ToArray();
        StoredFiles.WriteBoard(DraftBoardSource.PublishedPathOf(dataRoot, listed.Identity), board, rows);

        return rows.ToDictionary(row => row.Name, row => StoredFiles.PathOf(dataRoot, row), StringComparer.OrdinalIgnoreCase);
    }

    // DataManager pointed at an empty data root again.
    public static void Unload(string emptyFolder) => DataManager.LoadFrom(emptyFolder, "does-not-exist.json");
}
