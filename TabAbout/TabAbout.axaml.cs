using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace CRT
{
    public partial class TabAbout : UserControl
    {
        public TabAbout()
        {
            this.InitializeComponent();
        }

        // ###########################################################################################
        // Initializes static About-tab content (title/version/changelog) from assembly metadata.
        // ###########################################################################################
        public void InitializeAbout(Assembly assembly, string? versionString)
        {
            this.AboutAssemblyTitleText.Text = this.GetAssemblyTitle(assembly);
            this.AppVersionText.Text = versionString ?? "(unknown)";
            this.ChangelogTextBox.Text = this.LoadTextAsset("Assets/Changelog.txt");
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
        // Builds and displays a grouped credits list from the loaded board data.
        // ###########################################################################################
        private void PopulateCreditsSection(List<CreditEntry>? credits)
        {
            this.CreditsPanel.Children.Clear();

            if (credits == null || credits.Count == 0)
            {
                this.CreditsSectionBorder.IsVisible = false;
                return;
            }

            var categoryOrder = new List<string>();
            var subCategoryOrder = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var entriesByKey = new Dictionary<string, List<CreditEntry>>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in credits)
            {
                var cat = entry.Category;
                var sub = entry.SubCategory ?? string.Empty;
                var key = $"{cat}\u001F{sub}";

                if (!subCategoryOrder.ContainsKey(cat))
                {
                    categoryOrder.Add(cat);
                    subCategoryOrder[cat] = new List<string>();
                }

                if (!entriesByKey.ContainsKey(key))
                    subCategoryOrder[cat].Add(sub);

                if (!entriesByKey.TryGetValue(key, out var bucket))
                {
                    bucket = new List<CreditEntry>();
                    entriesByKey[key] = bucket;
                }

                bucket.Add(entry);
            }

            bool firstCategory = true;
            foreach (var category in categoryOrder)
            {
                this.CreditsPanel.Children.Add(new TextBlock
                {
                    Text = category,
                    FontWeight = FontWeight.SemiBold,
                    FontSize = 13,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, firstCategory ? 0 : 10, 0, 2)
                });
                firstCategory = false;

                foreach (var sub in subCategoryOrder[category])
                {
                    double nameIndent;

                    if (!string.IsNullOrWhiteSpace(sub))
                    {
                        this.CreditsPanel.Children.Add(new TextBlock
                        {
                            Text = sub,
                            FontSize = 12,
                            Opacity = 0.65,
                            TextWrapping = TextWrapping.Wrap,
                            Margin = new Thickness(14, 2, 0, 1)
                        });
                        nameIndent = 28;
                    }
                    else
                    {
                        nameIndent = 14;
                    }

                    var key = $"{category}\u001F{sub}";
                    if (!entriesByKey.TryGetValue(key, out var entries))
                        continue;

                    foreach (var entry in entries)
                        this.CreditsPanel.Children.Add(this.BuildCreditEntryRow(entry, nameIndent));
                }
            }

            this.CreditsSectionBorder.IsVisible = true;
        }

        // ###########################################################################################
        // Builds one credits row with optional clickable contact.
        // ###########################################################################################
        private Control BuildCreditEntryRow(CreditEntry entry, double nameIndent)
        {
            bool isClickable = !string.IsNullOrWhiteSpace(entry.Contact)
                && (IsContactWebUrl(entry.Contact) || IsContactEmail(entry.Contact));

            if (!isClickable)
            {
                var nameText = string.IsNullOrWhiteSpace(entry.Contact)
                    ? $"• {entry.NameOrHandle}"
                    : $"• {entry.NameOrHandle}  —  {entry.Contact}";

                return new TextBlock
                {
                    Text = nameText,
                    FontSize = 12,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(nameIndent, 1, 0, 1)
                };
            }

            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(nameIndent, 1, 0, 1)
            };

            row.Children.Add(new TextBlock
            {
                Text = $"• {entry.NameOrHandle}  —  ",
                FontSize = 12,
                VerticalAlignment = VerticalAlignment.Center
            });

            var href = BuildContactHref(entry.Contact);
            var linkButton = new Button
            {
                Content = entry.Contact,
                FontSize = 12,
                Padding = new Thickness(0),
                BorderThickness = new Thickness(0),
                VerticalAlignment = VerticalAlignment.Center
            };
            linkButton.Classes.Add("CreditsContactLink");
            linkButton.Click += (_, _) => this.OpenUrl(href);
            row.Children.Add(linkButton);

            return row;
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
}