using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// "THE stable source", "THE BETA source" (owner decision, 2026-10-01 - AppConfig.GetOnlineSourceLabel).
//
// The label is the bare name ("stable source"), so every sentence it goes into supplies its own
// "the". When the names changed, two callers were given one and four were not, and the sync banner
// read "Checking data from stable source" beside "Downloading file [1] of [9] from the stable
// source" (code review, 2026-10-01). Nothing type-checks a sentence, so this reads the source: every
// place the label is put into text must have "the " right before it.
// ###########################################################################################
public sealed class OnlineSourceLabelWordingTests
{
    private const string Call = "{AppConfig.GetOnlineSourceLabel()}";

    [Fact]
    public void Every_sentence_naming_the_data_source_says_the()
    {
        string root = OnlineSourceLabelWordingTests.ResolveAppFolder();
        var missing = new List<string>();

        foreach (string file in Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories))
        {
            string relative = Path.GetRelativePath(root, file);

            if (relative.Split(Path.DirectorySeparatorChar).Any(part => part is "bin" or "obj"))
                continue;

            string[] lines = File.ReadAllLines(file);

            for (int index = 0; index < lines.Length; index++)
            {
                foreach (Match match in Regex.Matches(lines[index], Regex.Escape(OnlineSourceLabelWordingTests.Call)))
                {
                    string before = lines[index][..match.Index];

                    if (!before.EndsWith("the ", StringComparison.Ordinal))
                        missing.Add($"{relative}:{index + 1}");
                }
            }
        }

        Assert.True(missing.Count == 0, "The source label without \"the\" before it at: " + string.Join(", ", missing));
    }

    private static string ResolveAppFolder()
    {
        string relative = Path.Combine("src", "CRT.App");
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative)))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Path.Combine(directory.FullName, relative);
    }
}
