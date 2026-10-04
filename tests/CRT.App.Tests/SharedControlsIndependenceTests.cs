using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** THE SHARED CONTROLS REACH INTO NOTHING OF CRT'S OWN ORCHESTRATION (code review,
// 2026-09-30). ***
//
// src/CRT.App/Controls/ - the table editor, BusyOverlay, ListRowDrag - was the separate CRT.UI
// library until 2026-09-30, whose rule was "no Main, no DataManager, no UserSettings, no Logger".
// The compiler enforced it then, because the library could not see CRT.App. Folded into CRT.App,
// a call to UserSettings from the table compiles fine - and it bypasses the host hand-over (a
// remembered choice passed in, reported back through FilterWantedChanged) that keeps headless
// tests of the table off the user's real settings file. CRT.Data's CrtLog is the logging seam.
//
// So this reads every .cs file under Controls/ with its comments and string literals removed
// (the files DISCUSS Main and DataManager in their comments, legitimately) and fails on any of the
// four names left in the code.
// ###########################################################################################
public sealed class SharedControlsIndependenceTests
{
    private static readonly Regex ForbiddenName =
        new(@"\b(?:Main|DataManager|UserSettings|Logger)\b", RegexOptions.CultureInvariant);

    // Comments and string/char literals, consumed LEFT TO RIGHT as one alternation, so a "//"
    // inside a string and a quote inside a comment are each read as what they really are.
    // Interpolated holes ($"{...}") are removed with their string - a known blind spot, accepted.
    private static readonly Regex CommentOrLiteral = new(
        @"//[^\n]*|/\*.*?\*/|@""(?:""""|[^""])*""|""(?:\\.|[^""\\\n])*""|'(?:\\.|[^'\\\n])+'",
        RegexOptions.Singleline | RegexOptions.CultureInvariant);

    // The forbidden names a piece of C# uses in CODE, in order of appearance.
    internal static IReadOnlyList<string> ForbiddenNamesIn(string source)
    {
        string codeOnly = CommentOrLiteral.Replace(source, " ");

        return ForbiddenName.Matches(codeOnly).Select(match => match.Value).ToList();
    }

    [Fact]
    public void No_file_under_Controls_uses_Main_DataManager_UserSettings_or_Logger()
    {
        string controls = SharedControlsIndependenceTests.ResolveControlsFolder();
        string[] files = Directory.GetFiles(controls, "*.cs", SearchOption.AllDirectories);

        // A folder that moved would otherwise make this pass on nothing.
        Assert.NotEmpty(files);

        var offences = files
            .SelectMany(file => SharedControlsIndependenceTests.ForbiddenNamesIn(File.ReadAllText(file))
                .Select(name => $"{Path.GetRelativePath(controls, file)}: {name}"))
            .ToList();

        Assert.True(offences.Count == 0,
            "The shared controls must be handed what they need by their host, not reach into CRT's " +
            "orchestration:" + Environment.NewLine + string.Join(Environment.NewLine, offences));
    }

    // ###########################################################################################
    // The scanner itself: a use in code is found, a mention in a comment or a string is not - and
    // neither a "//" inside a string nor a quote inside a comment hides the code after it.
    // ###########################################################################################
    [Theory]
    [InlineData("_ = UserSettings.EnableWorklog;", "UserSettings")]
    [InlineData("var window = (CRT.Main)topLevel;", "Main")]
    [InlineData("DataManager.ClearBoardCache();", "DataManager")]
    [InlineData("Logger.Info(\"x\");", "Logger")]
    [InlineData("// Main hands the table its choices", "")]
    [InlineData("/* Logger\n   and DataManager */ int a = 1;", "")]
    [InlineData("string s = \"UserSettings\";", "")]
    [InlineData("string s = @\"a \"\"Logger\"\" b\";", "")]
    [InlineData("string url = \"http://example.com\"; Logger.Info(url);", "Logger")]
    [InlineData("// it's \"quoted\"\nUserSettings.Save();", "UserSettings")]
    [InlineData("char c = '\"'; DataManager.Load();", "DataManager")]
    [InlineData("var mainWindow = MainWindow; CrtLog.Warning(\"x\");", "")]
    public void The_scanner_finds_names_in_code_and_ignores_comments_and_strings(string source, string expected)
    {
        Assert.Equal(expected, string.Join(",", SharedControlsIndependenceTests.ForbiddenNamesIn(source)));
    }

    private static string ResolveControlsFolder()
    {
        string relative = Path.Combine("src", "CRT.App", "Controls");
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative)))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Path.Combine(directory.FullName, relative);
    }
}
