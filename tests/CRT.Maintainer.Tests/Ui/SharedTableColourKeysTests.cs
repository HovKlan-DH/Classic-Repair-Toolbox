using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Styling;

namespace CRT.Maintainer.Tests.Ui;

// ###########################################################################################
// Every colour the shared table (src/CRT.UI) names resolves in THIS application, light and dark.
//
// The table borrows some of its colours from whichever application hosts it - Table_Bg and
// Table_BorderRowLine, the Button_Cancel_* red of "Save changes" - and CRT defines them in its
// own theme. A key named by the table and defined in CRT alone compiles, runs and draws nothing here: a missing
// DynamicResource is silently tolerated. CRT.App.Tests' SharedTableColourKeysTests asks the same
// of CRT, so a key added to one application only fails on the other side.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SharedTableColourKeysTests
{
    private static IReadOnlyList<string> KeysTheTableNames()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        string relative = Path.Combine("src", "CRT.UI", "BoardTable");

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative)))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Directory.GetFiles(Path.Combine(directory.FullName, relative), "*.axaml")
            .SelectMany(file => Regex.Matches(File.ReadAllText(file), @"\{DynamicResource\s+([A-Za-z0-9_]+)\s*\}").Select(match => match.Groups[1].Value))
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    [Fact]
    public void Every_colour_the_shared_table_names_resolves_in_both_themes()
    {
        IReadOnlyList<string> keys = KeysTheTableNames();

        // The parse found both kinds: a key the host lends, and one of the table's own.
        Assert.Contains("Table_Bg", keys);
        Assert.Contains("BoardTable_Header_Bg", keys);

        UiTest.Run(() =>
        {
            Application app = Application.Current!;

            foreach (ThemeVariant theme in new[] { ThemeVariant.Light, ThemeVariant.Dark })
            {
                List<string> missing = keys.Where(key => !app.TryGetResource(key, theme, out _)).ToList();

                Assert.True(missing.Count == 0, $"{theme}: {string.Join(", ", missing)}");
            }
        });
    }
}
