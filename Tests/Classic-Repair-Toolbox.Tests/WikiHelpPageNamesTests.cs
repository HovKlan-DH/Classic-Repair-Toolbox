using CRT;

namespace ClassicRepairToolbox.Tests;

// Pins the Wiki pages the shipped app's "?" help buttons open against the mirrored files in
// Assets/Wiki/.
//
// This is the test that stops the wrong-help-page bug coming back. A Wiki page rename is a pure
// content change - nothing in the build, the markup or the type system knows a URL string names a
// page - so renaming Workbooks.md and leaving the button's literal behind compiled clean, passed
// the whole suite, and simply landed users on the Wiki's front page. That is exactly what happened
// (the Configuration tab's Workbooks help button), and what CLAUDE.md warns about: renaming one of
// these pages breaks a button in builds already installed, which no update can fix.
//
// Rule 6 (no hardware, network, display or processes) is respected: the click itself goes through
// ExternalTargetLauncher's Process.Start and is deliberately untested - see
// ConfigurationHelpIconTests. What IS asserted is the page NAME the button was built with, which
// is plain string and file-existence work.
//
// Assets/Wiki/ is the source of truth for the published Wiki (nothing publishes automatically; the
// maintainer pastes each file across), so a page present here is the page that exists - or is
// about to - and a page absent here is one the button cannot reach.
public class WikiHelpPageNamesTests
{
    // Every Wiki page named by a help button in the shipped app, and where the button lives. Keep
    // this in step with AppConfig - a help button pointed at a page with no entry here is a page
    // nobody has verified exists.
    public static IEnumerable<object[]> InAppHelpPages() => new List<object[]>
    {
        new object[] { AppConfig.WikiPageWorkbooks, "Configuration tab, \"?\" beside \"Enable Workbooks tab\"" },
        new object[] { AppConfig.WikiPageMiniPro, "Configuration tab \"?\", and the component popup" },
        new object[] { AppConfig.WikiPageScopeKeyboard, "Component popup, numpad oscilloscope controls" },
        new object[] { AppConfig.WikiPageScopeSync, "Component popup, oscilloscope synchronization" },
    };

    // The whole point: a renamed or deleted page leaves the button opening a URL that resolves to
    // the Wiki front page, with nothing failing anywhere.
    [Theory]
    [MemberData(nameof(InAppHelpPages))]
    public void Every_in_app_help_button_names_a_page_that_exists_in_the_wiki_mirror(string pageName, string openedFrom)
    {
        string path = ResolveRepositoryPath($"Assets/Wiki/{pageName}.md");

        Assert.True(
            File.Exists(path),
            $"The help button in [{openedFrom}] opens Wiki page [{pageName}], but Assets/Wiki/{pageName}.md does not exist - " +
            "that button lands on the Wiki front page instead, in every build already installed");
    }

    // Filenames ARE page names on the GitHub Wiki, capitals included, so a page name carrying a
    // path, an extension or a URL-unsafe character cannot address the page it means to.
    [Theory]
    [MemberData(nameof(InAppHelpPages))]
    public void A_help_page_name_is_a_bare_page_name_with_no_path_or_extension(string pageName, string openedFrom)
    {
        Assert.False(string.IsNullOrWhiteSpace(pageName), $"[{openedFrom}] has no page name at all");
        Assert.DoesNotContain("/", pageName);
        Assert.DoesNotContain("\\", pageName);
        Assert.DoesNotContain(" ", pageName);
        Assert.False(pageName.EndsWith(".md", StringComparison.OrdinalIgnoreCase),
            $"[{openedFrom}] names [{pageName}] - the Wiki addresses pages without the extension");
    }

    // The URL builder is what removed five copies of the owner/repo literal, so it has to actually
    // produce the address those literals did.
    [Fact]
    public void The_wiki_url_is_built_from_the_repository_owner_and_name()
    {
        Assert.Equal(
            $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}/wiki/{AppConfig.WikiPageMiniPro}",
            AppConfig.WikiPageUrl(AppConfig.WikiPageMiniPro));
    }

    // ###########################################################################################
    // Walks up from the test binary until the repository file is found - the same approach
    // FontAwesomeAssetTests uses to reach the shipped Assets folder.
    // ###########################################################################################
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

        // Return a path that does not exist rather than throwing, so the assertion above reports
        // the missing PAGE rather than a resolution failure that reads like a broken test.
        return Path.Combine(AppContext.BaseDirectory, relativePath);
    }
}
