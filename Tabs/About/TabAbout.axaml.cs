using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Windows.Input;

namespace CRT
{
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
            if (!ExternalTargetLauncher.TryOpen(url))
            {
                Logger.Warning($"Rejected external target from About tab: [{url}]");
            }
        }

        // ###########################################################################################
        // Opens the GitHub project page from the About tab.
        // ###########################################################################################
        private void OnGitHubProjectPageClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://github.com/HovKlan-DH/Classic-Repair-Toolbox");
        }

        // ###########################################################################################
        // Opens the helper page from the About tab.
        // ###########################################################################################
        private void OnHelperPageClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://classic-repair-toolbox.dk");
        }

        // ###########################################################################################
        // Opens the YouTube channel from the About tab.
        // ###########################################################################################
        private void OnYouTubeChannelClick(object? sender, RoutedEventArgs e)
        {
            OpenUrl("https://www.youtube.com/@HovKlan-DH");
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