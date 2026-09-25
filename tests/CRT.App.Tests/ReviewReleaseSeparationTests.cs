using System.Text.RegularExpressions;
using CRT;

namespace ClassicRepairToolbox.Tests;

// Pins the two things that keep the review application's releases away from CRT's update check.
//
// Velopack's GithubSource reads the 10 most recent releases of the repository CRT names in
// AppConfig, merges every releases.<channel>.json it finds, and picks the highest version without
// looking at the package id (verified by decompiling Velopack 1.2.158 - the review app's prefixed
// packId does NOT separate them). So a review release that reached CRT's repository on CRT's
// default "win" or "linux" channel would put review-app packages in the update feed of every
// installed CRT, and a review version above an installed CRT's would be offered to it as an update.
//
// Two independent guards, each tested here:
//   1. The review release is published to its OWN repository, never the one AppConfig names.
//   2. The review app packs on its own "review-" channels, which CRT's packs never use - so even a
//      release that did land in CRT's repository would not be read by CRT.
//
// Nothing compiles or type-checks a workflow file, so these are the only things that notice either
// guard going missing. They read BOTH workflows and CRT's own AppConfig, because each rule has two
// sides and a test of one side alone passes happily while the two disagree.
public sealed class ReviewReleaseSeparationTests
{
    private const string ReviewWorkflow = ".github/workflows/build-and-release-review.yml";
    private const string CrtWorkflow = ".github/workflows/build-and-release.yml";

    [Fact]
    public void The_review_app_is_released_into_a_repository_CRT_does_not_update_from()
    {
        string[] lines = ReadWorkflow(ReviewWorkflow);
        string crtRepository = $"{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}";

        int releaseSteps = lines.Count(line => line.Contains("softprops/action-gh-release", StringComparison.Ordinal));
        List<string> targets = lines
            .Select(line => Regex.Match(line, @"^\s*repository:\s*(.+?)\s*$"))
            .Where(match => match.Success)
            .Select(match => ResolveEnv(lines, match.Groups[1].Value))
            .ToList();

        Assert.True(releaseSteps >= 1, $"No release step found in {ReviewWorkflow} - the parse has gone wrong");

        // A release step with no "repository:" publishes to the repository the workflow runs in,
        // which is CRT's.
        Assert.Equal(releaseSteps, targets.Count);

        foreach (string target in targets)
        {
            Assert.False(string.IsNullOrWhiteSpace(target), $"A release step in {ReviewWorkflow} names no repository");
            Assert.False(
                string.Equals(target, crtRepository, StringComparison.OrdinalIgnoreCase),
                $"{ReviewWorkflow} publishes into [{target}], the repository CRT checks for updates");
        }
    }

    [Fact]
    public void Every_review_app_package_is_packed_on_a_review_channel()
    {
        IReadOnlyList<string> packs = ReadPackCommands(ReviewWorkflow);

        // Windows and Linux. Fewer means the parse missed one, which would let this pass vacuously.
        Assert.True(packs.Count >= 2, $"Expected the Windows and Linux packs in {ReviewWorkflow}, found {packs.Count}");

        foreach (string pack in packs)
        {
            string? channel = ChannelOf(pack);
            Assert.True(
                channel != null && channel.StartsWith("review-", StringComparison.Ordinal),
                $"A review-app pack in {ReviewWorkflow} is on channel [{channel ?? "(default)"}] - without its own " +
                "\"review-\" channel its packages would join CRT's update feed:\n" + pack);
        }
    }

    [Fact]
    public void The_review_app_packs_each_platform_on_a_channel_of_its_own()
    {
        // Both packs are copied into one flat folder for the release, so two packs on one channel
        // would overwrite each other's releases.<channel>.json and leave one platform with no feed.
        List<string?> channels = ReadPackCommands(ReviewWorkflow).Select(ChannelOf).ToList();

        Assert.Equal(channels.Count, channels.Distinct().Count());
    }

    [Fact]
    public void No_CRT_package_is_packed_on_a_review_channel()
    {
        IReadOnlyList<string> packs = ReadPackCommands(CrtWorkflow);

        // Windows, Linux and the two macOS packs.
        Assert.True(packs.Count >= 4, $"Expected at least four packs in {CrtWorkflow}, found {packs.Count}");

        foreach (string pack in packs)
        {
            string? channel = ChannelOf(pack);
            Assert.False(
                channel != null && channel.StartsWith("review-", StringComparison.Ordinal),
                $"A CRT pack in {CrtWorkflow} is on review channel [{channel}]:\n" + pack);
        }
    }

    // "${{ env.NAME }}" resolved against the workflow's top-level "  NAME: value" line; anything
    // else is returned as written.
    private static string ResolveEnv(string[] lines, string value)
    {
        Match reference = Regex.Match(value, @"^\$\{\{\s*env\.([A-Za-z_][A-Za-z0-9_]*)\s*\}\}$");
        if (!reference.Success)
        {
            return value;
        }

        string name = reference.Groups[1].Value;
        Match definition = lines
            .Select(line => Regex.Match(line, $@"^\s*{name}:\s*(\S+)\s*$"))
            .FirstOrDefault(match => match.Success) ?? Match.Empty;

        return definition.Success ? definition.Groups[1].Value : string.Empty;
    }

    // Each "vpk pack" command with its continuation lines - a trailing backtick (PowerShell) or
    // backslash (bash) carries the command onto the next line.
    private static IReadOnlyList<string> ReadPackCommands(string relativePath)
    {
        string[] lines = ReadWorkflow(relativePath);
        var packs = new List<string>();

        for (int i = 0; i < lines.Length; i++)
        {
            if (!Regex.IsMatch(lines[i], @"\bvpk pack\b"))
            {
                continue;
            }

            var command = new List<string> { lines[i].Trim() };
            while (IsContinued(lines[i]) && i + 1 < lines.Length)
            {
                i++;
                command.Add(lines[i].Trim());
            }
            packs.Add(string.Join(Environment.NewLine, command));
        }

        return packs;
    }

    private static bool IsContinued(string line)
    {
        string trimmed = line.TrimEnd();
        return trimmed.EndsWith('`') || trimmed.EndsWith('\\');
    }

    // Null when the command names no channel, which means vpk's per-OS default ("win", "linux").
    private static string? ChannelOf(string pack)
    {
        Match match = Regex.Match(pack, @"(?:^|\s)(?:--channel|-c)\s+""?([^\s""`\\]+)", RegexOptions.Multiline);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string[] ReadWorkflow(string relativePath)
    {
        string path = ResolveRepositoryPath(relativePath);
        Assert.True(File.Exists(path), $"{relativePath} was not found above the test binaries");
        return File.ReadAllLines(path);
    }

    // Walks up from the test binary until the repository file is found - the same approach
    // WikiHelpPageNamesTests uses to reach Assets/Wiki.
    private static string ResolveRepositoryPath(string relativePath)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory != null)
        {
            string candidate = Path.Combine(directory.FullName, relativePath);
            if (File.Exists(candidate))
            {
                return candidate;
            }
            directory = directory.Parent;
        }

        return Path.Combine(AppContext.BaseDirectory, relativePath);
    }
}
