using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform;
using DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
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
            this.AboutAssemblyTitleText.Text = this.GetAssemblyTitle(assembly);
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
        // Updates the credits section for the currently loaded board.
        // ###########################################################################################
        public void SetCredits(List<CreditEntry>? credits)
        {
            this.PopulateCreditsSection(credits);
        }

        // ###########################################################################################
        // Resolves assembly title from metadata, with a fallback to assembly name.
        // ###########################################################################################
        private string GetAssemblyTitle(Assembly assembly)
        {
            var titleAttribute = assembly.GetCustomAttribute<AssemblyTitleAttribute>();
            if (!string.IsNullOrWhiteSpace(titleAttribute?.Title))
                return titleAttribute.Title;

            return assembly.GetName().Name ?? "Classic Repair Toolbox";
        }

        // ###########################################################################################
        // Loads a text asset from Avalonia resources and returns the raw file content.
        // ###########################################################################################
        private string LoadTextAsset(string assetPath)
        {
            try
            {
                var assetUri = new Uri($"avares://Classic-Repair-Toolbox/{assetPath}");
                using var stream = AssetLoader.Open(assetUri);
                using var reader = new StreamReader(stream);
                return reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load changelog [{ex.Message}]");
                return "Unable to load changelog...";
            }
        }

        // ###########################################################################################
        // Opens the configured URL in the system default browser.
        // ###########################################################################################
        private void OpenUrl(string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open URL - [{url}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Opens the GitHub project page from the About tab.
        // ###########################################################################################
        private void OnGitHubProjectPageClick(object? sender, RoutedEventArgs e)
        {
            this.OpenUrl("https://github.com/HovKlan-DH/Classic-Repair-Toolbox");
        }

        // ###########################################################################################
        // Opens the helper page from the About tab.
        // ###########################################################################################
        private void OnHelperPageClick(object? sender, RoutedEventArgs e)
        {
            this.OpenUrl("https://classic-repair-toolbox.dk");
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
                    && (IsContactWebUrl(entry.Contact) || IsContactEmail(entry.Contact));

                Action? openAction = null;
                if (isClickable && !string.IsNullOrWhiteSpace(entry.Contact))
                {
                    string href = BuildContactHref(entry.Contact);
                    openAction = () => this.OpenUrl(href);
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

        // ###########################################################################################
        // Returns true when the contact string looks like a web URL.
        // ###########################################################################################
        private static bool IsContactWebUrl(string contact)
            => contact.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || contact.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            || contact.StartsWith("www.", StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Returns true when the contact string looks like an email address.
        // ###########################################################################################
        private static bool IsContactEmail(string contact)
            => contact.Contains('@') && !contact.Contains(' ');

        // ###########################################################################################
        // Builds the href to open from contact text.
        // ###########################################################################################
        private static string BuildContactHref(string contact)
        {
            if (IsContactEmail(contact))
                return $"mailto:{contact}";
            if (contact.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
                return $"https://{contact}";
            return contact;
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