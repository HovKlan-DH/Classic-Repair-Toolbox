using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Styling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Every colour the shared table (src/CRT.UI) names resolves in THIS application, light and dark.
//
// The table borrows some of its colours from whichever application hosts it - Table_Bg and
// Table_BorderRowLine, the Button_Cancel_* red of "Save changes". A key named by the table and
// defined nowhere here compiles, runs and draws nothing: a missing DynamicResource is silently
// tolerated.
//
// The same is asked of the Maintainer tab's own markup (2026-09-29): it was the separate CRT
// Maintainer application, whose MaintainerApp.axaml defined its colours; they live in CRT's
// App.axaml now, and a key left behind there would draw nothing here.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SharedTableColourKeysTests
{
    private static IReadOnlyList<string> KeysTheTableNames() =>
        SharedTableColourKeysTests.KeysNamedIn(Path.Combine("src", "CRT.UI", "BoardTable"));

    private static IReadOnlyList<string> KeysTheMaintainerTabNames() =>
        SharedTableColourKeysTests.KeysNamedIn(Path.Combine("src", "CRT.App", "Tabs", "Maintainer"));

    // Every {DynamicResource X} in the .axaml files of one repository folder.
    private static IReadOnlyList<string> KeysNamedIn(string relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative)))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Directory.GetFiles(Path.Combine(directory.FullName, relative), "*.axaml")
            .SelectMany(file => Regex.Matches(File.ReadAllText(file), @"\{DynamicResource\s+([A-Za-z0-9_]+)\s*\}").Select(match => match.Groups[1].Value))
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    private static void AssertResolvesInBothThemes(IReadOnlyList<string> keys) =>
        UiTest.Run(() =>
        {
            Application app = Application.Current!;

            foreach (ThemeVariant theme in new[] { ThemeVariant.Light, ThemeVariant.Dark })
            {
                List<string> missing = keys.Where(key => !app.TryGetResource(key, theme, out _)).ToList();

                Assert.True(missing.Count == 0, $"{theme}: {string.Join(", ", missing)}");
            }
        });

    // ###########################################################################################
    // *** THE MAINTAINER TAB'S OWN MARKUP (2026-09-29). *** "Approve and publish to BETA" is green
    // through Button_Ok_*, the mode buttons and their badges have colours of their own - all ten of
    // which moved from the separate application's MaintainerApp.axaml into CRT's App.axaml. Fails
    // if one is dropped, or if the tab ever names a colour CRT does not define.
    // ###########################################################################################
    [Fact]
    public void Every_colour_the_maintainer_tab_names_resolves_in_both_themes()
    {
        IReadOnlyList<string> keys = KeysTheMaintainerTabNames();

        // A parse that found nothing cannot pass: the decision buttons' colours, and two of the
        // keys that came over from the separate application.
        Assert.Contains("Button_Ok_Bg", keys);
        Assert.Contains("Button_Cancel_Bg", keys);
        Assert.Contains("Badge_Attention_Bg", keys);
        Assert.Contains("Mode_Selected_Bg", keys);

        SharedTableColourKeysTests.AssertResolvesInBothThemes(keys);
    }

    [Fact]
    public void Every_colour_the_shared_table_names_resolves_in_both_themes()
    {
        IReadOnlyList<string> keys = KeysTheTableNames();

        // The parse found both kinds: a key the host lends, and one of the table's own.
        Assert.Contains("Table_Bg", keys);
        Assert.Contains("BoardTable_Header_Bg", keys);

        SharedTableColourKeysTests.AssertResolvesInBothThemes(keys);
    }
}
