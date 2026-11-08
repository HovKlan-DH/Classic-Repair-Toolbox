using Avalonia;
using Avalonia.Controls.Templates;
using Handlers.DataHandling;
using Handlers.Theming;
using System.IO;
using System.Text;

namespace CRT
{
    // ###########################################################################################
    // TabAbout.Changelog.cs - THE CHANGELOG UNDER THE CREDITS (owner request, 2026-11-06: "In
    // 'About' and below the credits - can we then add the entire changelog and then style it alike
    // GitHub, so it gets indented and bold/italic etc.?").
    //
    // The text is the repository's CHANGELOG.md, embedded by CRT.App.csproj as it was when this
    // build was made - so it is always the changelog of the version running, and nothing is fetched.
    // GitHubMarkdown reads it, ChangelogVersionHeadings rewords each version's heading, MarkdownRenderer
    // draws it.
    //
    // EVERY BLOCK IS AN ITEM OF THE TAB'S VIRTUALIZED LIST, built as it scrolls into view (measured
    // in the render harness with real text, 2026-11-06: drawn at once it was some 1,100 controls -
    // 0.2 s to build and 0.25-0.45 s to lay out the first time, and 0.15-0.2 s to lay out again at
    // every change of the window's width). The items are made the first time the tab is shown, as
    // most runs never open it; a tab not shown is only detached, so that happens once for the run.
    // ###########################################################################################
    public partial class TabAbout
    {
        // The name CRT.App.csproj gives the embedded CHANGELOG.md.
        internal const string ChangelogResourceName = "CRT.CHANGELOG.md";

        private bool thisChangelogShown;

        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);
            this.ShowChangelog();
        }

        // Puts the changelog's blocks under the page, once. Internal so a test can ask for it
        // without a window.
        internal void ShowChangelog()
        {
            if (this.thisChangelogShown)
            {
                return;
            }

            this.thisChangelogShown = true;

            string? markdown = ReadEmbeddedChangelog();
            if (string.IsNullOrWhiteSpace(markdown))
            {
                // Never expected - the build embeds it - but a missing changelog is no reason to
                // show an empty heading.
                Logger.Warning($"The embedded changelog [{ChangelogResourceName}] was not found - the \"About\" tab shows none");
                this.ChangelogTitle.IsVisible = false;
                return;
            }

            // A block that leaves the view gives its place to the next one, and in between the list
            // hands the template no block at all - that builds nothing (it threw, 2026-11-06).
            var renderer = new MarkdownRenderer(this.FontSize, OpenUrl);
            this.AboutItems.ItemTemplate = new FuncDataTemplate<MarkdownBlockItem?>((item, _) => item == null ? null : renderer.Build(item));

            // Each version's heading "Version 3.0.0, 2026-October-11", the number bold (ChangelogVersionHeadings).
            foreach (var item in MarkdownRenderer.Arrange(ChangelogVersionHeadings.Of(GitHubMarkdown.Parse(markdown))))
            {
                this.AboutItems.Items.Add(item);
            }
        }

        // The CHANGELOG.md this build carries, or null when it carries none.
        internal static string? ReadEmbeddedChangelog()
        {
            using Stream? stream = typeof(TabAbout).Assembly.GetManifestResourceStream(ChangelogResourceName);
            if (stream == null)
            {
                return null;
            }

            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }
    }
}
