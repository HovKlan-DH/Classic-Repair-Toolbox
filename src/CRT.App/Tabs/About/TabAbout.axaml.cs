using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Windows.Input;

namespace CRT
{
    // ###########################################################################################
    // The "About" tab: CRT's name and version, its links, the board's revision date and credits,
    // and the changelog.
    //
    //   TabAbout.axaml.cs     - construction, version, links, the board's credits, Home/End/Page keys
    //   TabAbout.Changelog.cs - the changelog under the credits, each block an item of the tab's
    //                           virtualized list, built as it scrolls into view
    // ###########################################################################################
    public partial class TabAbout : UserControl
    {
        public ObservableCollection<CreditDisplayItem> CreditsList { get; } = new ObservableCollection<CreditDisplayItem>();

        public TabAbout()
        {
            this.InitializeComponent();
            this.CreditsItemsControl.ItemsSource = this.CreditsList;
        }

        // ###########################################################################################
        // Initializes static About-tab content (title/version/changelog) from assembly metadata.
        // ###########################################################################################
        public void InitializeAbout(Assembly assembly, string? versionString)
        {
            this.AboutAssemblyTitleText.Text = GetAssemblyTitle(assembly);
            this.AppVersionText.Text = versionString ?? "(unknown)";
        }

        // ###########################################################################################
        // Updates the board-specific information (revision date and credits).
        // ###########################################################################################
        public void SetBoardInfo(string? revisionDate, List<CreditEntry>? credits)
        {
            if (string.IsNullOrWhiteSpace(revisionDate))
            {
                this.RevisionDatePanel.IsVisible = false;
            }
            else
            {
                this.RevisionDateText.Text = revisionDate;
                this.RevisionDatePanel.IsVisible = true;
            }

            this.PopulateCreditsSection(credits);
        }

        // ###########################################################################################
        // HOME, END, PAGE UP AND PAGE DOWN SCROLL THE PAGE (owner request, 2026-11-06: "When I am placed anywhere
        // in the 'About' tab and I press 'Home', then I want it to navigate to the top of the page. Likewise with
        // 'End' should go to the end. Also, 'Page/Up' and down should work"). Main hands them over from wherever
        // the keyboard is (Main.KeyIsForAboutPage) - mostly the component filter, which takes the focus back at
        // every click on this tab. True when the key was taken.
        //
        // End brings the LAST BLOCK into view before going to the end: the list builds its blocks only as they
        // come into view and guesses the height of the rest, so the end worked out from that guess fell short.
        // ###########################################################################################
        internal bool TakeScrollKey(Key key)
        {
            if (this.AboutItems.FindDescendantOfType<ScrollViewer>() is not { } viewer)
            {
                return false;
            }

            switch (key)
            {
                case Key.Home:
                    viewer.ScrollToHome();
                    return true;

                case Key.End:
                    if (this.AboutItems.ItemCount > 0)
                    {
                        this.AboutItems.ScrollIntoView(this.AboutItems.ItemCount - 1);
                        this.AboutItems.UpdateLayout();
                    }

                    viewer.ScrollToEnd();
                    return true;

                case Key.PageUp:
                    viewer.PageUp();
                    return true;

                case Key.PageDown:
                    viewer.PageDown();
                    return true;

                default:
                    return false;
            }
        }

        // ###########################################################################################
        // Resolves assembly title from metadata, with a fallback to assembly name.
        // ###########################################################################################
        private static string GetAssemblyTitle(Assembly assembly)
        {
            var titleAttribute = assembly.GetCustomAttribute<AssemblyTitleAttribute>();
            if (!string.IsNullOrWhiteSpace(titleAttribute?.Title))
                return titleAttribute.Title;

            return assembly.GetName().Name ?? "Classic Repair Toolbox";
        }

        // ###########################################################################################
        // Opens a validated external target through the shared launcher.
        // ###########################################################################################
        private static void OpenUrl(string url)
        {
            ExternalTargetLauncher.OpenOrSay(url);
        }

        // ###########################################################################################
        // Opens the GitHub project page from the About tab.
        // ###########################################################################################
        private void OnGitHubProjectPageClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://github.com/HovKlan-DH/Classic-Repair-Toolbox");
        }

        // ###########################################################################################
        // Opens the fun-facts page from the About tab.
        // ###########################################################################################
        private void OnFunFactsPageClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://classic-repair-toolbox.dk/funfacts/");
        }

        // ###########################################################################################
        // Opens the YouTube channel from the About tab.
        // ###########################################################################################
        private void OnYouTubeChannelClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://www.youtube.com/@HovKlan-DH");
        }

        // ###########################################################################################
        // Opens the invite to the developer's Discord server from the About tab.
        // ###########################################################################################
        private void OnDiscordServerClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://discord.gg/kVTtdvZtzE");
        }

        // ###########################################################################################
        // Builds and displays a tabular credits list from the loaded board data.
        // ###########################################################################################
        private void PopulateCreditsSection(List<CreditEntry>? credits)
        {
            this.CreditsList.Clear();

            if (credits == null || credits.Count == 0)
            {
                this.CreditsSectionBorder.IsVisible = false;
                return;
            }

            foreach (var entry in credits)
            {
                bool isClickable = !string.IsNullOrWhiteSpace(entry.Contact)
                    && (ContactLinkFormatter.IsContactWebUrl(entry.Contact) || ContactLinkFormatter.IsContactEmail(entry.Contact));

                Action? openAction = null;
                if (isClickable && !string.IsNullOrWhiteSpace(entry.Contact))
                {
                    string href = ContactLinkFormatter.BuildContactHref(entry.Contact);
                    openAction = () => OpenUrl(href);
                }

                this.CreditsList.Add(new CreditDisplayItem(
                    entry.Category,
                    entry.SubCategory ?? string.Empty,
                    entry.NameOrHandle,
                    entry.Contact ?? string.Empty,
                    isClickable,
                    openAction
                ));
            }

            this.CreditsSectionBorder.IsVisible = true;
        }
    }

    public class CreditDisplayItem
    {
        public string Category { get; }
        public string SubCategory { get; }
        public string Name { get; }
        public string Contact { get; }
        public bool IsLink { get; }
        public ICommand? OpenContactCommand { get; }

        public CreditDisplayItem(string category, string subCategory, string name, string contact, bool isLink, Action? openAction)
        {
            this.Category = category;
            this.SubCategory = subCategory;
            this.Name = name;
            this.Contact = contact;
            this.IsLink = isLink;
            if (openAction != null)
            {
                this.OpenContactCommand = new ActionCommand(openAction);
            }
        }
    }
}