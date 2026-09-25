using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** NO TEST HARD-CODES A WINDOWS PATH (owner request, 2026-09-25). ***
//
// The suite is written and run on Windows; GitHub runs it on Linux (ubuntu-latest), where a
// backslash is not a path separator. So a literal like @"C:\Drafts\Commodore\.crt-draft.json" is a
// full path on the developer's machine and ONE long file name on the runner: the test passes
// locally, the local test run that gates every change passes, and it fails only after the push.
// That happened three or four times - most recently DraftFolderLayoutTests, which held back every
// build and both release workflows - while the rule against it lived only in notes nobody is made
// to read at the moment of writing a test.
//
// This makes the rule a failing test, which fails on WINDOWS too, so it is caught where the test
// is written. Build a path with Path.Combine (from Path.GetTempPath() when it must be rooted). The
// rare literal that is safe on purpose - a hostile input the code must refuse everywhere, command-
// line text compared as text, a test that returns early off Windows - carries a comment on its
// line or within the three lines above it:
//
//     // windows-path-literal: <why it is safe on Linux>
//
// A line guarded by OperatingSystem.IsWindows() in the same window needs no comment.
// ###########################################################################################
public sealed class TestPathLiteralTests
{
    private const string Marker = "windows-path-literal:";

    // A drive letter, a colon and a backslash, not glued to a word before it: "C:\x", @"C:\x",
    // "--root=D:\x". Found anywhere on a line, since a path can sit inside a longer string.
    private static readonly Regex WindowsPath = new(@"(?<![A-Za-z0-9_])[A-Za-z]:\\", RegexOptions.Compiled);

    // Lines above a literal that may carry its marker or its OS guard.
    private const int Window = 3;

    [Fact]
    public void No_test_hard_codes_a_Windows_path_without_saying_why_it_is_safe_on_Linux()
    {
        string testsRoot = TestPathLiteralTests.RepositoryPath("tests");
        var offenders = new List<string>();

        foreach (string file in Directory.EnumerateFiles(testsRoot, "*.cs", SearchOption.AllDirectories))
        {
            string relative = Path.GetRelativePath(testsRoot, file).Replace('\\', '/');

            // Build output, and this file, whose own samples below are the thing being detected.
            if (relative.Contains("/bin/", StringComparison.Ordinal) ||
                relative.Contains("/obj/", StringComparison.Ordinal) ||
                relative.EndsWith("/" + nameof(TestPathLiteralTests) + ".cs", StringComparison.Ordinal))
            {
                continue;
            }

            string[] lines = File.ReadAllLines(file);

            for (int index = 0; index < lines.Length; index++)
            {
                if (TestPathLiteralTests.IsUnexplained(lines, index))
                    offenders.Add($"tests/{relative}:{index + 1}: {lines[index].Trim()}");
            }
        }

        Assert.True(
            offenders.Count == 0,
            "A test hard-codes a Windows path, which is one long FILE NAME on the Linux CI runner - it passes " +
            "here and fails only after the push. Build it with Path.Combine; if it is safe on Linux on purpose, " +
            $"say why in a \"// {TestPathLiteralTests.Marker} ...\" comment on or just above the line:\n" +
            string.Join("\n", offenders));
    }

    // ---- The detector itself, so a broken pattern cannot let everything through ------------------

    [Theory]
    [MemberData(nameof(Caught))]
    public void A_Windows_path_literal_is_caught(string line)
    {
        Assert.True(TestPathLiteralTests.IsUnexplained([line], 0), line);
    }

    [Theory]
    [MemberData(nameof(NotCaught))]
    public void A_line_without_one_or_with_a_comment_or_guard_is_not_caught(string[] lines)
    {
        Assert.False(TestPathLiteralTests.IsUnexplained(lines, lines.Length - 1), string.Join(" | ", lines));
    }

    public static TheoryData<string> Caught() => new()
    {
        """    private const string Root = @"C:\Drafts";""",
        """    string root = "C:\\Drafts";""",
        """    Resolve(new[] { @"--data-root=D:\somewhere\Data" });""",
        """    [InlineData("c:\\windows\\evil.dll")]""",
    };

    public static TheoryData<string[]> NotCaught() => new()
    {
        new[] { """    string root = Path.Combine(Path.GetTempPath(), "Drafts");""" },
        new[] { """    string url = "https://example.com/a:b";""" },
        new[] { """    // never a literal @"C:\..." - see above""" },
        new[] { $"""    [InlineData("C:\\Windows\\evil.dll")] // {TestPathLiteralTests.Marker} a hostile input""" },
        new[] { $"""    // {TestPathLiteralTests.Marker} command-line text""", """    Assert.Equal(""", """        @"D:\somewhere\Data",""" },
        new[] { """    private static string Root => OperatingSystem.IsWindows()""", """        ? @"C:\Users\test\Data" """ },
    };

    // ###########################################################################################
    // Whether line `index` holds a Windows path literal with no marker and no OS guard on it or in
    // the Window lines above. A comment line is never an offender - it may describe the rule.
    // ###########################################################################################
    private static bool IsUnexplained(IReadOnlyList<string> lines, int index)
    {
        string line = lines[index];

        if (line.TrimStart().StartsWith("//", StringComparison.Ordinal) || !TestPathLiteralTests.WindowsPath.IsMatch(line))
            return false;

        for (int above = Math.Max(0, index - TestPathLiteralTests.Window); above <= index; above++)
        {
            if (lines[above].Contains(TestPathLiteralTests.Marker, StringComparison.Ordinal) ||
                lines[above].Contains("OperatingSystem.IsWindows()", StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
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
