using System.Text.RegularExpressions;

namespace CRT.Data.Tests;

// ###########################################################################################
// *** EVERY BOARD BUILT BY HAND NAMES ITS CAPTION. ***
//
// A BoardData is copied by listing its sections one by one, and the "# Hardware:" / "# Board:"
// caption is the part those lists forget. It has been lost that way THREE times: a new board's
// draft (DraftSeeder), the Schematic images window (BoardFilesWindow), and on 2026-09-28 the
// publish itself plus "Save to draft" and the label editor's save - every sheet of a workbook
// approved into BETA came out with its first two lines empty (owner report).
//
// So this reads the source: every `new BoardData { ... }` (and BoardData's own `=> new() { ... }`
// copies) must assign HardwareName AND BoardName, even when the answer is string.Empty - which
// makes whoever adds the next one decide, rather than forget.
// ###########################################################################################
public sealed class BoardDataCaptionTests
{
    [Fact]
    public void Every_BoardData_built_by_hand_names_its_caption()
    {
        var missing = new List<string>();
        int seen = 0;

        foreach (string file in Directory.EnumerateFiles(BoardDataCaptionTests.RepositoryPath("src"), "*.cs", SearchOption.AllDirectories))
        {
            string normalised = file.Replace('\\', '/');

            if (normalised.Contains("/bin/", StringComparison.Ordinal) || normalised.Contains("/obj/", StringComparison.Ordinal))
                continue;

            string text = File.ReadAllText(file);
            bool isBoardData = Path.GetFileName(file) == "BoardData.cs";

            foreach (Match match in BoardDataCaptionTests.Starts(isBoardData).Matches(text))
            {
                int open = match.Index + match.Length - 1;
                string initializer = BoardDataCaptionTests.Block(text, open);
                seen++;

                if (!initializer.Contains("HardwareName =", StringComparison.Ordinal) ||
                    !initializer.Contains("BoardName =", StringComparison.Ordinal))
                {
                    int line = text[..match.Index].Count(character => character == '\n') + 1;
                    missing.Add($"{Path.GetRelativePath(BoardDataCaptionTests.RepositoryPath(""), file)}:{line}");
                }
            }
        }

        // The scan found the copies it is meant to guard - a pattern matching nothing would pass
        // for ever.
        Assert.True(seen >= 10, $"Only {seen} board initialisers found - has the pattern stopped matching?");
        Assert.True(missing.Count == 0, "A BoardData built without its caption (HardwareName / BoardName): " + string.Join(", ", missing));
    }

    // `new BoardData {` anywhere; in BoardData.cs also its `=> new() {` copy methods, whose type is
    // the method's own.
    private static Regex Starts(bool isBoardData) => isBoardData
        ? new Regex(@"(?:new\s+BoardData|BoardData\s+\w+\([^)]*\)\s*=>\s*new\s*\(\s*\))\s*\{", RegexOptions.Compiled)
        : new Regex(@"new\s+BoardData\s*(?:\(\s*\))?\s*\{", RegexOptions.Compiled);

    // The initializer from its opening brace to the matching closing one.
    private static string Block(string text, int open)
    {
        int depth = 0;

        for (int index = open; index < text.Length; index++)
        {
            if (text[index] == '{')
                depth++;
            else if (text[index] == '}' && --depth == 0)
                return text[open..(index + 1)];
        }

        return text[open..];
    }

    // Walks up from the test binary to the repository root (the folder holding the solution).
    private static string RepositoryPath(string relative)
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, relative);
    }
}
