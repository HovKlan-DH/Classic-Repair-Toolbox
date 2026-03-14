using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CRT
{
    public sealed class ContributionComponentRow
    {
        public string BoardLabel { get; set; } = string.Empty;
        public string FriendlyName { get; set; } = string.Empty;
        public string TechnicalNameOrValue { get; set; } = string.Empty;
        public string PartNumber { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Region { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public sealed class ContributionComponentImageRow
    {
        public string BoardLabel { get; set; } = string.Empty;
        public string Region { get; set; } = string.Empty;
        public string Pin { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string ExpectedOscilloscopeReading { get; set; } = string.Empty;
        public string File { get; set; } = string.Empty;
        public string Note { get; set; } = string.Empty;
        public Bitmap? PreviewImage { get; set; }
        public string PreviewStatusText { get; set; } = "No preview available";
    }

    public sealed class ContributionComponentLocalFileRow
    {
        public string BoardLabel { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string File { get; set; } = string.Empty;
    }

    public sealed class ContributionComponentLinkRow
    {
        public string BoardLabel { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    public sealed class ContributionBoardLocalFileRow
    {
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string File { get; set; } = string.Empty;
    }

    public sealed class ContributionBoardLinkRow
    {
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    public sealed class ComponentContributionPayload
    {
        public string ApplicationVersion { get; set; } = string.Empty;
        public string HardwareName { get; set; } = string.Empty;
        public string BoardName { get; set; } = string.Empty;
        public string Region { get; set; } = string.Empty;
        public string ComponentBoardLabel { get; set; } = string.Empty;
        public string ComponentDisplayText { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public DateTimeOffset SubmittedUtc { get; set; }
        public List<ContributionComponentRow> Components { get; set; } = new();
        public List<ContributionComponentImageRow> ComponentImages { get; set; } = new();
        public List<ComponentHighlightEntry> ComponentHighlights { get; set; } = new();
        public List<ContributionComponentLocalFileRow> ComponentLocalFiles { get; set; } = new();
        public List<ContributionComponentLinkRow> ComponentLinks { get; set; } = new();
        public List<ContributionBoardLocalFileRow> BoardLocalFiles { get; set; } = new();
        public List<ContributionBoardLinkRow> BoardLinks { get; set; } = new();
    }

    public partial class ComponentContributionWindow : Window
    {
        private readonly ObservableCollection<ContributionComponentRow> thisComponentRows = new();
        private readonly ObservableCollection<ContributionComponentImageRow> thisComponentImageRows = new();
        private readonly ObservableCollection<ContributionComponentLocalFileRow> thisComponentLocalFileRows = new();
        private readonly ObservableCollection<ContributionComponentLinkRow> thisComponentLinkRows = new();
        private readonly ObservableCollection<ContributionBoardLocalFileRow> thisBoardLocalFileRows = new();
        private readonly ObservableCollection<ContributionBoardLinkRow> thisBoardLinkRows = new();
        private readonly List<ComponentHighlightEntry> thisComponentHighlightRows = new();

        private string thisHardwareName = string.Empty;
        private string thisBoardName = string.Empty;
        private string thisLocalRegion = string.Empty;
        private string thisBoardLabel = string.Empty;
        private string thisComponentDisplayText = string.Empty;
        private string thisDataRoot = string.Empty;

        public ComponentContributionWindow()
        {
            this.InitializeComponent();

            this.ComponentRowsItemsControl.ItemsSource = this.thisComponentRows;
            this.ComponentImageRowsItemsControl.ItemsSource = this.thisComponentImageRows;
            this.ComponentLocalFileRowsItemsControl.ItemsSource = this.thisComponentLocalFileRows;
            this.ComponentLinkRowsItemsControl.ItemsSource = this.thisComponentLinkRows;
            this.BoardLocalFileRowsItemsControl.ItemsSource = this.thisBoardLocalFileRows;
            this.BoardLinkRowsItemsControl.ItemsSource = this.thisBoardLinkRows;

            this.EmailTextBox.Text = UserSettings.ContactEmail;
            this.Closed += this.OnWindowClosed;

            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Loads the selected component context and all editable rows into the window.
        // ###########################################################################################
        public void LoadComponent(BoardData boardData, string dataRoot, string hardwareName, string boardName, string region, string boardLabel)
        {
            this.thisDataRoot = dataRoot;
            this.thisHardwareName = hardwareName;
            this.thisBoardName = boardName;
            this.thisLocalRegion = region;
            this.thisBoardLabel = boardLabel;

            var primaryComponent = boardData.Components.FirstOrDefault(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase));

            this.thisComponentDisplayText = this.BuildComponentDisplayText(primaryComponent, boardLabel);

            this.PopulateHeader();
            this.LoadRows(boardData, boardLabel);
        }

        // ###########################################################################################
        // Populates the window header with the selected component and board context.
        // ###########################################################################################
        private void PopulateHeader()
        {
            this.Title = string.IsNullOrWhiteSpace(this.thisBoardLabel)
                ? "Component contribution"
                : $"Component contribution - {this.thisBoardLabel}";

            this.ComponentTitleTextBlock.Text = this.thisComponentDisplayText;
            this.HardwareContextTextBlock.Text = $"Hardware: {this.thisHardwareName}";
            this.BoardContextTextBlock.Text = $"Board: {this.thisBoardName}";
            this.RegionContextTextBlock.Text = $"Region: {this.thisLocalRegion}";
            this.ComponentImagesRegionTextBlock.Text = $"Component images relevant for the {this.thisLocalRegion} region";
        }

        // ###########################################################################################
        // Loads editable row collections from the selected board and component.
        // ###########################################################################################
        private void LoadRows(BoardData boardData, string boardLabel)
        {
            foreach (var row in this.thisComponentImageRows)
            {
                this.DisposeComponentImagePreview(row);
            }

            this.thisComponentRows.Clear();
            this.thisComponentImageRows.Clear();
            this.thisComponentLocalFileRows.Clear();
            this.thisComponentLinkRows.Clear();
            this.thisBoardLocalFileRows.Clear();
            this.thisBoardLinkRows.Clear();
            this.thisComponentHighlightRows.Clear();

            foreach (var row in boardData.Components.Where(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase)))
            {
                this.thisComponentRows.Add(new ContributionComponentRow
                {
                    BoardLabel = row.BoardLabel,
                    FriendlyName = row.FriendlyName,
                    TechnicalNameOrValue = row.TechnicalNameOrValue,
                    PartNumber = row.PartNumber,
                    Category = row.Category,
                    Region = row.Region,
                    Description = row.Description
                });
            }

            foreach (var row in boardData.ComponentImages.Where(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase) &&
                (string.IsNullOrWhiteSpace(c.Region) ||
                 string.Equals(c.Region.Trim(), this.thisLocalRegion, StringComparison.OrdinalIgnoreCase))))
            {
                this.thisComponentImageRows.Add(new ContributionComponentImageRow
                {
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    Pin = row.Pin,
                    Name = row.Name,
                    ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                    File = row.File,
                    Note = row.Note
                });
            }

            foreach (var row in boardData.ComponentHighlights.Where(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase)))
            {
                this.thisComponentHighlightRows.Add(new ComponentHighlightEntry
                {
                    SchematicName = row.SchematicName,
                    BoardLabel = row.BoardLabel,
                    X = row.X,
                    Y = row.Y,
                    Width = row.Width,
                    Height = row.Height
                });
            }

            foreach (var row in boardData.ComponentLocalFiles.Where(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase)))
            {
                this.thisComponentLocalFileRows.Add(new ContributionComponentLocalFileRow
                {
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    File = row.File
                });
            }

            foreach (var row in boardData.ComponentLinks.Where(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase)))
            {
                this.thisComponentLinkRows.Add(new ContributionComponentLinkRow
                {
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    Url = row.Url
                });
            }

            foreach (var row in boardData.BoardLocalFiles)
            {
                this.thisBoardLocalFileRows.Add(new ContributionBoardLocalFileRow
                {
                    Category = row.Category,
                    Name = row.Name,
                    File = row.File
                });
            }

            foreach (var row in boardData.BoardLinks)
            {
                this.thisBoardLinkRows.Add(new ContributionBoardLinkRow
                {
                    Category = row.Category,
                    Name = row.Name,
                    Url = row.Url
                });
            }

            this.RefreshAllComponentImagePreviews();
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Builds a compact display label for the selected component.
        // ###########################################################################################
        private string BuildComponentDisplayText(ComponentEntry? component, string boardLabel)
        {
            if (component == null)
            {
                return boardLabel;
            }

            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(component.BoardLabel))
                parts.Add(component.BoardLabel.Trim());
            if (!string.IsNullOrWhiteSpace(component.FriendlyName))
                parts.Add(component.FriendlyName.Trim());
            if (!string.IsNullOrWhiteSpace(component.TechnicalNameOrValue))
                parts.Add(component.TechnicalNameOrValue.Trim());

            return parts.Count == 0 ? boardLabel : string.Join(" | ", parts);
        }

/*
        // ###########################################################################################
        // Adds a new editable row to the Components section.
        // ###########################################################################################
        private void OnAddComponentRowClick(object? sender, RoutedEventArgs e)
        {
            this.thisComponentRows.Add(new ContributionComponentRow
            {
                BoardLabel = this.thisBoardLabel,
                Region = this.thisLocalRegion
            });
        }
*/

        // ###########################################################################################
        // Removes an editable row from the Components section.
        // ###########################################################################################
        private void OnRemoveComponentRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentRow row })
            {
                this.thisComponentRows.Remove(row);
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Component images section.
        // ###########################################################################################
        private void OnAddComponentImageRowClick(object? sender, RoutedEventArgs e)
        {
            var row = new ContributionComponentImageRow
            {
                BoardLabel = this.thisBoardLabel,
                Region = this.thisLocalRegion
            };

            InsertRowAtTop(this.thisComponentImageRows, row);
            this.RefreshComponentImagePreview(row);
            this.RefreshComponentImageRowsItemsControl();
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Component images section.
        // ###########################################################################################
        private void OnRemoveComponentImageRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentImageRow row })
            {
                this.DisposeComponentImagePreview(row);
                this.thisComponentImageRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Component local files section.
        // ###########################################################################################
        private void OnAddComponentLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            InsertRowAtTop(this.thisComponentLocalFileRows, new ContributionComponentLocalFileRow
            {
                BoardLabel = this.thisBoardLabel
            });

            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Component local files section.
        // ###########################################################################################
        private void OnRemoveComponentLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentLocalFileRow row })
            {
                this.thisComponentLocalFileRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Component links section.
        // ###########################################################################################
        private void OnAddComponentLinkRowClick(object? sender, RoutedEventArgs e)
        {
            InsertRowAtTop(this.thisComponentLinkRows, new ContributionComponentLinkRow
            {
                BoardLabel = this.thisBoardLabel
            });

            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Component links section.
        // ###########################################################################################
        private void OnRemoveComponentLinkRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentLinkRow row })
            {
                this.thisComponentLinkRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Board local files section.
        // ###########################################################################################
        private void OnAddBoardLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            InsertRowAtTop(this.thisBoardLocalFileRows, new ContributionBoardLocalFileRow());
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Board local files section.
        // ###########################################################################################
        private void OnRemoveBoardLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionBoardLocalFileRow row })
            {
                this.thisBoardLocalFileRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Board links section.
        // ###########################################################################################
        private void OnAddBoardLinkRowClick(object? sender, RoutedEventArgs e)
        {
            InsertRowAtTop(this.thisBoardLinkRows, new ContributionBoardLinkRow());
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Board links section.
        // ###########################################################################################
        private void OnRemoveBoardLinkRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionBoardLinkRow row })
            {
                this.thisBoardLinkRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Closes the window without sending anything.
        // ###########################################################################################
        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // ###########################################################################################
        // Validates the edited payload and submits it to the contribution backend.
        // ###########################################################################################
        private async void OnSubmitClick(object? sender, RoutedEventArgs e)
        {
            string email = this.EmailTextBox.Text?.Trim() ?? string.Empty;
            string comment = this.MandatoryCommentTextBox.Text?.Trim() ?? string.Empty;

            if (!string.IsNullOrEmpty(email) && !Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                this.ShowStatus("Please enter a valid email address", true);
                return;
            }

            if (string.IsNullOrWhiteSpace(comment))
            {
                this.ShowStatus("Please provide a mandatory change comment before sending", true);
                return;
            }

            UserSettings.ContactEmail = email;
            this.SubmitButton.IsEnabled = false;

            try
            {
                IProgress<string> progress = new Progress<string>(statusMessage =>
                {
                    this.ShowStatus(statusMessage, false);
                });

                var result = await this.ProcessAndSendContributionAsync(email, comment, progress);

                if (result.Success)
                {
                    this.ShowStatus("Contribution submitted successfully - thank you :-)", false);
                }
                else
                {
                    Logger.Warning($"Component contribution submission failed. HTTP {result.StatusCode}. Server responded with: {result.ResponseBody}");

                    if (result.StatusCode == 404)
                    {
                        this.ShowStatus("Failed to send contribution: Server endpoint not found (HTTP 404)", true);
                    }
                    else
                    {
                        this.ShowStatus($"Failed to send contribution (HTTP {result.StatusCode}) - please check the logfile for details", true);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Exception while sending component contribution: {ex}");
                this.ShowStatus("Network or system error while sending contribution - please try again later...", true);
            }
            finally
            {
                this.SubmitButton.IsEnabled = true;
            }
        }

        // ###########################################################################################
        // Builds the contribution payload, zips it together with any referenced files, and posts it.
        // ###########################################################################################
        private async Task<(bool Success, int StatusCode, string ResponseBody)> ProcessAndSendContributionAsync(string email, string comment, IProgress<string> progress)
        {
            progress.Report("Preparing contribution payload...");

            var payload = this.BuildPayload(email, comment);
            string payloadJson = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            using var memoryStream = new MemoryStream();

            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
            {
                this.AddTextEntryToZip(archive, "ComponentContribution.json", payloadJson);

                var referencedFiles = this.GatherExistingReferencedFiles();
                for (int i = 0; i < referencedFiles.Count; i++)
                {
                    progress.Report($"Packaging referenced files... {i + 1}/{referencedFiles.Count}");
                    this.AddFileToZipSafe(archive, referencedFiles[i].Source, referencedFiles[i].ZipEntryName);
                }
            }

            memoryStream.Position = 0;

            using var httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(1)
            };

            using var formContent = new MultipartFormDataContent();
            formContent.Add(new StringContent(email), "email");
            formContent.Add(new StringContent(this.BuildContributionFeedbackText(comment)), "feedback");
            formContent.Add(new StringContent(AppConfig.AppVersionString), "version");
            formContent.Add(new StringContent("component-contribution"), "submissionType");

            var fileContent = new ByteArrayContent(memoryStream.ToArray());
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/zip");
            formContent.Add(fileContent, "attachmentFile", "ComponentContributionPayload.zip");

            using var progressContent = new ProgressableStreamContent(formContent, percent =>
                progress.Report($"Sending to server... {percent}%"));

            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(AppConfig.AppShortName + " " + AppConfig.AppVersionString);

            var response = await httpClient.PostAsync("https://classic-repair-toolbox.dk/app-contribution/", progressContent);
            string responseBody = await response.Content.ReadAsStringAsync();
            bool isSuccess = response.IsSuccessStatusCode &&
                             responseBody.Trim().StartsWith("Success", StringComparison.OrdinalIgnoreCase);

            return (isSuccess, (int)response.StatusCode, responseBody);
        }

        // ###########################################################################################
        // Builds a structured payload representing the current edited state of the window.
        // ###########################################################################################
        private ComponentContributionPayload BuildPayload(string email, string comment)
        {
            return new ComponentContributionPayload
            {
                ApplicationVersion = AppConfig.AppVersionString,
                HardwareName = this.thisHardwareName,
                BoardName = this.thisBoardName,
                Region = this.thisLocalRegion,
                ComponentBoardLabel = this.thisBoardLabel,
                ComponentDisplayText = this.thisComponentDisplayText,
                Email = email,
                Comment = comment,
                SubmittedUtc = DateTimeOffset.UtcNow,
                Components = this.thisComponentRows.Select(row => new ContributionComponentRow
                {
                    BoardLabel = row.BoardLabel?.Trim() ?? string.Empty,
                    FriendlyName = row.FriendlyName?.Trim() ?? string.Empty,
                    TechnicalNameOrValue = row.TechnicalNameOrValue?.Trim() ?? string.Empty,
                    PartNumber = row.PartNumber?.Trim() ?? string.Empty,
                    Category = row.Category?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                    Description = row.Description?.Trim() ?? string.Empty
                }).ToList(),
                ComponentImages = this.thisComponentImageRows.Select(row => new ContributionComponentImageRow
                {
                    BoardLabel = row.BoardLabel?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                    Pin = row.Pin?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty,
                    Note = row.Note?.Trim() ?? string.Empty
                }).ToList(),
                ComponentHighlights = this.thisComponentHighlightRows.Select(row => new ComponentHighlightEntry
                {
                    SchematicName = row.SchematicName?.Trim() ?? string.Empty,
                    BoardLabel = row.BoardLabel?.Trim() ?? string.Empty,
                    X = row.X?.Trim() ?? string.Empty,
                    Y = row.Y?.Trim() ?? string.Empty,
                    Width = row.Width?.Trim() ?? string.Empty,
                    Height = row.Height?.Trim() ?? string.Empty
                }).ToList(),
                ComponentLocalFiles = this.thisComponentLocalFileRows.Select(row => new ContributionComponentLocalFileRow
                {
                    BoardLabel = row.BoardLabel?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty
                }).ToList(),
                ComponentLinks = this.thisComponentLinkRows.Select(row => new ContributionComponentLinkRow
                {
                    BoardLabel = row.BoardLabel?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty
                }).ToList(),
                BoardLocalFiles = this.thisBoardLocalFileRows.Select(row => new ContributionBoardLocalFileRow
                {
                    Category = row.Category?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty
                }).ToList(),
                BoardLinks = this.thisBoardLinkRows.Select(row => new ContributionBoardLinkRow
                {
                    Category = row.Category?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty
                }).ToList()
            };
        }

        // ###########################################################################################
        // Builds the plain-text feedback summary sent alongside the zipped JSON payload.
        // ###########################################################################################
        private string BuildContributionFeedbackText(string comment)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Component contribution submission");
            builder.AppendLine($"Hardware: {this.thisHardwareName}");
            builder.AppendLine($"Board: {this.thisBoardName}");
            builder.AppendLine($"Component: {this.thisComponentDisplayText}");
            builder.AppendLine($"Region context: {this.thisLocalRegion}");
            builder.AppendLine();
            builder.AppendLine("Mandatory change comment:");
            builder.AppendLine(comment);

            return builder.ToString();
        }

        // ###########################################################################################
        // Resolves existing local files referenced by edited rows so they can be attached as context.
        // ###########################################################################################
        private List<(string Source, string ZipEntryName)> GatherExistingReferencedFiles()
        {
            var files = new List<(string Source, string ZipEntryName)>();
            var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            this.AddResolvedFilesFromPaths(
                this.thisComponentImageRows.Select(r => r.File),
                "ReferencedFiles/ComponentImages",
                files,
                seenPaths);

            this.AddResolvedFilesFromPaths(
                this.thisComponentLocalFileRows.Select(r => r.File),
                "ReferencedFiles/ComponentLocalFiles",
                files,
                seenPaths);

            this.AddResolvedFilesFromPaths(
                this.thisBoardLocalFileRows.Select(r => r.File),
                "ReferencedFiles/BoardLocalFiles",
                files,
                seenPaths);

            return files;
        }

        // ###########################################################################################
        // Adds resolved file references from edited rows to the outgoing zip attachment list.
        // ###########################################################################################
        private void AddResolvedFilesFromPaths(IEnumerable<string> pathValues, string zipFolder, List<(string Source, string ZipEntryName)> files, HashSet<string> seenPaths)
        {
            int index = files.Count;

            foreach (var pathValue in pathValues)
            {
                var resolvedPath = this.ResolveExistingFilePath(pathValue);
                if (string.IsNullOrWhiteSpace(resolvedPath) || !seenPaths.Add(resolvedPath))
                {
                    continue;
                }

                index++;
                string zipEntryName = $"{zipFolder}/{index:D3}_{Path.GetFileName(resolvedPath)}";
                files.Add((resolvedPath, zipEntryName));
            }
        }

        // ###########################################################################################
        // Resolves an edited file path either as absolute or relative to the current data root.
        // ###########################################################################################
        private string? ResolveExistingFilePath(string pathValue)
        {
            string trimmed = pathValue?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return null;
            }

            if (Path.IsPathRooted(trimmed) && File.Exists(trimmed))
            {
                return trimmed;
            }

            string relativeToDataRoot = Path.Combine(this.thisDataRoot, trimmed.Replace('/', Path.DirectorySeparatorChar));
            return File.Exists(relativeToDataRoot) ? relativeToDataRoot : null;
        }

        // ###########################################################################################
        // Adds a UTF-8 text file entry to the output zip archive.
        // ###########################################################################################
        private void AddTextEntryToZip(ZipArchive archive, string entryName, string content)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
            writer.Write(content);
        }

        // ###########################################################################################
        // Adds a file to the provided zip archive, skipping unreadable files safely.
        // ###########################################################################################
        private void AddFileToZipSafe(ZipArchive archive, string sourcePath, string entryName)
        {
            if (!File.Exists(sourcePath))
            {
                return;
            }

            try
            {
                using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                using var entryStream = entry.Open();
                fs.CopyTo(entryStream);
            }
            catch
            {
                // Ignore unreadable files.
            }
        }

        // ###########################################################################################
        // Updates the status text on the UI thread using success or error styling.
        // ###########################################################################################
        private void ShowStatus(string message, bool isError)
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.StatusTextBlock.Text = message;

                if (isError)
                {
                    this.StatusTextBlock.Classes.Add("error");
                    this.StatusTextBlock.Classes.Remove("success");
                }
                else
                {
                    this.StatusTextBlock.Classes.Add("success");
                    this.StatusTextBlock.Classes.Remove("error");
                }

                this.StatusTextBlock.IsVisible = true;
            });
        }

        // ###########################################################################################
        // Refreshes all component image previews from the currently edited file paths.
        // ###########################################################################################
        private void RefreshAllComponentImagePreviews()
        {
            foreach (var row in this.thisComponentImageRows)
            {
                this.RefreshComponentImagePreview(row);
            }

            this.RefreshComponentImageRowsItemsControl();
        }

        // ###########################################################################################
        // Refreshes a single component image preview from its current file path.
        // ###########################################################################################
        private void RefreshComponentImagePreview(ContributionComponentImageRow row)
        {
            this.DisposeComponentImagePreview(row);

            string? resolvedPath = this.ResolveExistingFilePath(row.File);
            if (string.IsNullOrWhiteSpace(resolvedPath))
            {
                row.PreviewStatusText = string.Empty;
                return;
            }

            try
            {
                row.PreviewImage = new Bitmap(resolvedPath);
                row.PreviewStatusText = string.Empty;
            }
            catch (Exception ex)
            {
                row.PreviewStatusText = string.Empty;
                Logger.Warning($"Failed to load contribution image preview [{resolvedPath}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Disposes the currently loaded preview image for a component image row.
        // ###########################################################################################
        private void DisposeComponentImagePreview(ContributionComponentImageRow row)
        {
            row.PreviewImage?.Dispose();
            row.PreviewImage = null;
        }

        // ###########################################################################################
        // Rebinds the component image item source so preview updates are reflected immediately.
        // ###########################################################################################
        private void RefreshComponentImageRowsItemsControl()
        {
            this.ComponentImageRowsItemsControl.ItemsSource = null;
            this.ComponentImageRowsItemsControl.ItemsSource = this.thisComponentImageRows;
        }

        // ###########################################################################################
        // Refreshes an image preview when the edited file path box loses focus.
        // ###########################################################################################
        private void OnComponentImageFileTextBoxLostFocus(object? sender, RoutedEventArgs e)
        {
            if (sender is TextBox { Tag: ContributionComponentImageRow row })
            {
                this.RefreshComponentImagePreview(row);
                this.RefreshComponentImageRowsItemsControl();
            }
        }

        // ###########################################################################################
        // Disposes loaded preview images when the contribution window closes.
        // ###########################################################################################
        private void OnWindowClosed(object? sender, EventArgs e)
        {
            foreach (var row in this.thisComponentImageRows)
            {
                this.DisposeComponentImagePreview(row);
            }
        }

        // ###########################################################################################
        // Opens a file picker for any file-backed row and applies the selected path.
        // ###########################################################################################
        private async void OnFileTextBoxPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            e.Handled = true;

            if (sender is not TextBox textBox)
            {
                return;
            }

            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null)
            {
                return;
            }

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Select file",
                AllowMultiple = false
            });

            if (files == null || files.Count == 0)
            {
                return;
            }

            string selectedPath = files[0].Path.LocalPath;
            textBox.Text = selectedPath;
            this.ApplySelectedFilePath(textBox.Tag, selectedPath);
        }

        // ###########################################################################################
        // Applies a newly selected file path to the corresponding row model.
        // ###########################################################################################
        private void ApplySelectedFilePath(object? tag, string selectedPath)
        {
            switch (tag)
            {
                case ContributionComponentImageRow componentImageRow:
                    componentImageRow.File = selectedPath;
                    this.RefreshComponentImagePreview(componentImageRow);
                    this.RefreshComponentImageRowsItemsControl();
                    break;

                case ContributionComponentLocalFileRow componentLocalFileRow:
                    componentLocalFileRow.File = selectedPath;
                    break;

                case ContributionBoardLocalFileRow boardLocalFileRow:
                    boardLocalFileRow.File = selectedPath;
                    break;
            }
        }

        // ###########################################################################################
        // Opens a file picker for any file-backed row and applies the selected path.
        // ###########################################################################################
        private async void OnFileTextBoxPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            e.Handled = true;

            if (sender is not TextBox textBox)
            {
                return;
            }

            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null)
            {
                return;
            }

            string? currentPath = this.GetCurrentFilePath(textBox.Tag);
            string? suggestedStartLocation = this.GetSuggestedStartLocation(currentPath);

            var options = new FilePickerOpenOptions
            {
                Title = "Select file",
                AllowMultiple = false
            };

            if (!string.IsNullOrWhiteSpace(suggestedStartLocation) && Directory.Exists(suggestedStartLocation))
            {
                try
                {
                    options.SuggestedStartLocation = await topLevel.StorageProvider.TryGetFolderFromPathAsync(suggestedStartLocation);
                }
                catch
                {
                }
            }

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(options);
            if (files == null || files.Count == 0)
            {
                return;
            }

            string selectedPath = files[0].Path.LocalPath;
            textBox.Text = selectedPath;
            this.ApplySelectedFilePath(textBox.Tag, selectedPath);
        }

        // ###########################################################################################
        // Returns the current file path for the given tagged row object.
        // ###########################################################################################
        private string? GetCurrentFilePath(object? tag)
        {
            return tag switch
            {
                ContributionComponentImageRow componentImageRow => componentImageRow.File,
                ContributionComponentLocalFileRow componentLocalFileRow => componentLocalFileRow.File,
                ContributionBoardLocalFileRow boardLocalFileRow => boardLocalFileRow.File,
                _ => null
            };
        }

        // ###########################################################################################
        // Computes the best starting directory for the file picker based on the current file value.
        // ###########################################################################################
        private string? GetSuggestedStartLocation(string? currentPath)
        {
            string trimmed = currentPath?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return Directory.Exists(this.thisDataRoot) ? this.thisDataRoot : null;
            }

            if (Path.IsPathRooted(trimmed))
            {
                string? rootedDirectory = Path.GetDirectoryName(trimmed);
                if (!string.IsNullOrWhiteSpace(rootedDirectory) && Directory.Exists(rootedDirectory))
                {
                    return rootedDirectory;
                }
            }

            string combinedPath = Path.Combine(this.thisDataRoot, trimmed.Replace('/', Path.DirectorySeparatorChar));
            string? combinedDirectory = Path.GetDirectoryName(combinedPath);
            if (!string.IsNullOrWhiteSpace(combinedDirectory) && Directory.Exists(combinedDirectory))
            {
                return combinedDirectory;
            }

            return Directory.Exists(this.thisDataRoot) ? this.thisDataRoot : null;
        }

        // ###########################################################################################
        // Persists the shared email address when the field loses focus and the value is valid.
        // ###########################################################################################
        private void OnEmailTextBoxLostFocus(object? sender, RoutedEventArgs e)
        {
            string email = this.EmailTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(email) || Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                UserSettings.ContactEmail = email;
            }
        }

        // ###########################################################################################
        // Inserts a new row at the top of a collection so it becomes visible immediately.
        // ###########################################################################################
        private static void InsertRowAtTop<T>(ObservableCollection<T> collection, T row)
        {
            collection.Insert(0, row);
        }

        // ###########################################################################################
        // Scrolls the contribution editor to the top of the main content area.
        // ###########################################################################################
        private void OnScrollToTopClick(object? sender, RoutedEventArgs e)
        {
            this.MainScrollViewer.Offset = new Vector(this.MainScrollViewer.Offset.X, 0);
        }

        // ###########################################################################################
        // Scrolls the contribution editor to the bottom of the main content area.
        // ###########################################################################################
        private void OnScrollToBottomClick(object? sender, RoutedEventArgs e)
        {
            double bottomOffset = Math.Max(0, this.MainScrollViewer.Extent.Height - this.MainScrollViewer.Viewport.Height);
            this.MainScrollViewer.Offset = new Vector(this.MainScrollViewer.Offset.X, bottomOffset);
        }

        // ###########################################################################################
        // Updates the visible row counters for the editable contribution sections.
        // ###########################################################################################
        private void UpdateSectionCounters()
        {
            this.ComponentImagesCountTextBlock.Text = $"({this.thisComponentImageRows.Count})";
            this.ComponentLocalFilesCountTextBlock.Text = $"({this.thisComponentLocalFileRows.Count})";
            this.ComponentLinksCountTextBlock.Text = $"({this.thisComponentLinkRows.Count})";
            this.BoardLocalFilesCountTextBlock.Text = $"({this.thisBoardLocalFileRows.Count})";
            this.BoardLinksCountTextBlock.Text = $"({this.thisBoardLinkRows.Count})";
        }

    }
}