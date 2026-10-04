using System.Collections.Concurrent;
using System.IO;
using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** A TEST'S BOARD WORKBOOK, WRITTEN ONCE PER BOARD (2026-10-03, owner request: "so it can go
// even faster"). *** Most UI tests begin by writing a draft workbook, and many write the SAME
// board - "U1 and U2", "one component" - over and over. Each real write builds nine sheets in
// EPPlus, zips them and renames the file into place: about 13 s of a two-minute CRT.App.Tests run
// (measured with a sampled trace). BoardWorkbookWriter's output depends on the board alone - it
// is deterministic by design (BoardWorkbookWriterTests' byte-identical test), so the bytes of the
// first write of a board ARE the bytes of every later one, and this hands them out again.
//
// FOR A TEST'S SET-UP ONLY. A save the code under test makes still goes through the real writer -
// and BoardWorkbookWriter's own tests never come here. The key is the whole board as JSON, fields
// included, so two boards differing anywhere get their own bytes; CachedWorkbooksTests holds that
// for every property of every entry type.
// ###########################################################################################
internal static class CachedWorkbooks
{
    private static readonly JsonSerializerOptions KeyOptions = new() { IncludeFields = true };

    private static readonly ConcurrentDictionary<string, byte[]> BytesByBoard = new(StringComparer.Ordinal);

    // Writes `board` to `path` exactly as BoardWorkbookWriter.Write would, creating the folder.
    public static void Write(string path, BoardData board)
    {
        byte[] bytes = BytesByBoard.GetOrAdd(CachedWorkbooks.KeyOf(board), _ => CachedWorkbooks.WriteForReal(board));

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        File.WriteAllBytes(path, bytes);
    }

    internal static string KeyOf(BoardData board) => JsonSerializer.Serialize(board, KeyOptions);

    private static byte[] WriteForReal(BoardData board)
    {
        string temporary = Path.Combine(Path.GetTempPath(), $"crt-cached-workbook-{Guid.NewGuid():N}.xlsx");

        try
        {
            BoardWorkbookWriter.Write(temporary, board);
            return File.ReadAllBytes(temporary);
        }
        finally
        {
            File.Delete(temporary);
        }
    }
}
