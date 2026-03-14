using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CRT
{
    public partial class TabFeedback : UserControl
    {
        private readonly ObservableCollection<string> _customAttachments = new();

        public TabFeedback()
        {
            this.InitializeComponent();
            this.AttachmentsListBox.ItemsSource = this._customAttachments;
            this.EmailTextBox.Text = UserSettings.ContactEmail;
            this.EmailTextBox.LostFocus += this.OnEmailTextBoxLostFocus;

            this._customAttachments.CollectionChanged += (s, e) =>
            {
                this.ClearAttachmentsButton.IsEnabled = this._customAttachments.Count > 0;
            };
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
        // Prompts the user to select one or more files to include in the submission.
        // ###########################################################################################
        private async void OnAttachFilesClick(object? sender, RoutedEventArgs e)
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null) return;

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Select files to attach",
                AllowMultiple = true
            });

            foreach (var file in files)
            {
                if (!this._customAttachments.Contains(file.Path.LocalPath))
                {
                    this._customAttachments.Add(file.Path.LocalPath);
                }
            }
        }

        // ###########################################################################################
        // Prompts the user to select an entire directory to recursively attach.
        // ###########################################################################################
        private async void OnAttachFolderClick(object? sender, RoutedEventArgs e)
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null) return;

            var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Select a folder to attach",
                AllowMultiple = false
            });

            if (folders != null && folders.Count > 0)
            {
                string path = folders[0].Path.LocalPath;
                if (!this._customAttachments.Contains(path))
                {
                    this._customAttachments.Add(path);
                }
            }
        }

        // ###########################################################################################
        // Clears all custom selected attachments from the list.
        // ###########################################################################################
        private void OnClearAttachmentsClick(object? sender, RoutedEventArgs e)
        {
            this._customAttachments.Clear();
        }

        // ###########################################################################################
        // Removes a single custom attachment from the list.
        // ###########################################################################################
        private void OnRemoveAttachmentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string pathToRemove)
            {
                this._customAttachments.Remove(pathToRemove);
            }
        }

        // ###########################################################################################
        // Validates input, builds the zip archive, and dispatches the HTTP request.
        // ###########################################################################################
        private async void OnSubmitClick(object? sender, RoutedEventArgs e)
        {
            string email = this.EmailTextBox.Text?.Trim() ?? string.Empty;
            string feedback = this.FeedbackTextBox.Text?.Trim() ?? string.Empty;

            if (!string.IsNullOrEmpty(email) && !Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                this.ShowStatus("Please enter a valid email address", isError: true);
                return;
            }

            UserSettings.ContactEmail = email;

            if (string.IsNullOrEmpty(feedback))
            {
                this.ShowStatus("Please provide a description of your issue or suggestion before sending", isError: true);
                return;
            }

            this.SubmitButton.IsEnabled = false;

            IProgress<string> progress = new Progress<string>(statusMessage =>
            {
                this.ShowStatus(statusMessage, isError: false);
            });

            progress.Report("Preparing payload...");

            bool attachLogs = this.AttachLogfileCheckBox.IsChecked == true;
            bool attachConfig = this.AttachConfigsCheckBox.IsChecked == true;
            var customPaths = this._customAttachments.ToList();

            try
            {
                var (success, statusCode, responseBody) = await Task.Run(() =>
                    this.ProcessAndSendFeedbackAsync(email, feedback, attachLogs, attachConfig, customPaths, progress));

                if (success)
                {
                    this.ShowStatus("Feedback submitted successfully - thank you :-)", isError: false);
                    this.FeedbackTextBox.Text = string.Empty;
                    this._customAttachments.Clear();
                    this.AttachLogfileCheckBox.IsChecked = false;
                    this.AttachConfigsCheckBox.IsChecked = false;
                }
                else
                {
                    Logger.Warning($"Feedback submission failed. HTTP {(int)statusCode}. Server responded with: {responseBody}");

                    if (statusCode == 404)
                    {
                        this.ShowStatus("Failed to send feedback: Server endpoint not found (HTTP 404)", isError: true);
                    }
                    else
                    {
                        this.ShowStatus($"Failed to send feedback (HTTP {(int)statusCode}) - please check the logfile for details", isError: true);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Exception while sending feedback: {ex}");
                this.ShowStatus("Network or system error while sending feedback - please try again later...", isError: true);
            }
            finally
            {
                this.SubmitButton.IsEnabled = true;
            }
        }

        // ###########################################################################################
        // Helper to update the UI status text block from anywhere safely.
        // ###########################################################################################
        private void ShowStatus(string message, bool isError)
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.StatusTextBlock.Text = message;

                // Toggle the pseudo-classes to let XAML styles handle the color
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
        // Collects local files, generates the zip stream, and performs the multipart POST request.
        // ###########################################################################################
        private async Task<(bool Success, int StatusCode, string ResponseBody)> ProcessAndSendFeedbackAsync(string email, string feedbackText, bool attachLogs, bool attachConfig, List<string> customPaths, IProgress<string> progress)
        {
            var targetFiles = new List<(string Source, string ZipEntryName)>();
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string localAppFolder = Path.Combine(appData, AppConfig.AppFolderName);

            // 1. Gather all files to be zipped
            if (attachLogs)
            {
                targetFiles.Add((Path.Combine(localAppFolder, AppConfig.LogFileName), AppConfig.LogFileName));
            }

            if (attachConfig)
            {
                targetFiles.Add((Path.Combine(localAppFolder, AppConfig.SettingsFileName), AppConfig.SettingsFileName));
                targetFiles.Add((Path.Combine(localAppFolder, AppConfig.TracesFileName), AppConfig.TracesFileName));
            }

            foreach (string path in customPaths)
            {
                if (File.Exists(path))
                {
                    targetFiles.Add((path, Path.GetFileName(path)));
                }
                else if (Directory.Exists(path))
                {
                    string folderName = new DirectoryInfo(path).Name;
                    foreach (string filePath in Directory.GetFiles(path, "*.*", SearchOption.AllDirectories))
                    {
                        string relativePath = Path.GetRelativePath(path, filePath);
                        string zipEntryName = Path.Combine(folderName, relativePath).Replace('\\', '/');
                        targetFiles.Add((filePath, zipEntryName));
                    }
                }
            }

            // 2. Count the total raw size of files for precise Zipping progress
            long totalUncompressedBytes = 0;
            foreach (var file in targetFiles)
            {
                if (File.Exists(file.Source))
                {
                    try { totalUncompressedBytes += new FileInfo(file.Source).Length; } catch { }
                }
            }

            // 3. Zip files into memory stream
            using var memoryStream = new MemoryStream();
            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
            {
                long currentUncompressedBytes = 0;
                int lastReportedPercent = -1;

                if (totalUncompressedBytes == 0)
                {
                    progress.Report("Packaging payload... 100%");
                }

                foreach (var file in targetFiles)
                {
                    this.AddFileToZipSafe(archive, file.Source, file.ZipEntryName, bytesAdded =>
                    {
                        currentUncompressedBytes += bytesAdded;
                        if (totalUncompressedBytes > 0)
                        {
                            int percent = (int)((currentUncompressedBytes * 100) / totalUncompressedBytes);
                            if (percent != lastReportedPercent)
                            {
                                lastReportedPercent = percent;
                                progress.Report($"Packaging payload... {percent}%");
                            }
                        }
                    });
                }
            }

            memoryStream.Position = 0;

            // 4. Construct form payload 
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(1) };
            using var formContent = new MultipartFormDataContent();

            formContent.Add(new StringContent(email), "email");
            formContent.Add(new StringContent(feedbackText), "feedback");
            formContent.Add(new StringContent(AppConfig.AppVersionString), "version");

            if (memoryStream.Length > 22) // More than empty zip header
            {
                var fileContent = new ByteArrayContent(memoryStream.ToArray());
                fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/zip");
                formContent.Add(fileContent, "attachmentFile", "FeedbackPayload.zip");
            }

            // Track Upload progress
            using var progressContent = new ProgressableStreamContent(formContent, percent => progress.Report($"Sending to server... {percent}%"));

            //Target URL
            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("CRT "+ AppConfig.AppVersionString);
            var response = await httpClient.PostAsync("https://classic-repair-toolbox.dk/app-feedback/", progressContent);

            // Read the exact string back from the server
            // We must explicitly look for the string "Success" to evaluate a true success,
            // because PHP can crash with a warning string while technically returning HTTP 200
            string responseBody = await response.Content.ReadAsStringAsync();
            bool isSuccess = response.IsSuccessStatusCode && responseBody.Trim().StartsWith("Success", StringComparison.OrdinalIgnoreCase);

            return (isSuccess, (int)response.StatusCode, responseBody);
        }

        // ###########################################################################################
        // Adds a local file path to the provided ZipArchive, silencing read-access violations 
        // if file doesn't exist or is locked. Reports bytes copied back out to optionally track 
        // compression progress dynamically.
        // ###########################################################################################
        private void AddFileToZipSafe(ZipArchive archive, string sourcePath, string entryName, Action<int>? onBytesRead = null)
        {
            if (!File.Exists(sourcePath)) return;
            try
            {
                using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                using var entryStream = entry.Open();

                if (onBytesRead != null)
                {
                    var buffer = new byte[81920]; // 80 KB chunks
                    int bytesRead;
                    while ((bytesRead = fs.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        entryStream.Write(buffer, 0, bytesRead);
                        onBytesRead(bytesRead);
                    }
                }
                else
                {
                    fs.CopyTo(entryStream);
                }
            }
            catch
            {
                // Ignore files that are heavily locked or otherwise unreadable
            }
        }
    }

    // ###########################################################################################
    // A custom HttpContent wrapper to track the upload progress of an inner HttpContent payload.
    // ###########################################################################################
    public class ProgressableStreamContent : HttpContent
    {
        private readonly HttpContent _innerContent;
        private readonly Action<int> _progress;

        public ProgressableStreamContent(HttpContent innerContent, Action<int> progress)
        {
            this._innerContent = innerContent;
            this._progress = progress;

            foreach (var header in innerContent.Headers)
            {
                this.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
        }

        protected override async Task SerializeToStreamAsync(Stream stream, System.Net.TransportContext? context)
        {
            using var innerStream = new MemoryStream();
            await this._innerContent.CopyToAsync(innerStream);
            innerStream.Position = 0;

            var buffer = new byte[81920]; // 80 KB
            var totalLength = innerStream.Length;
            long uploadedBytes = 0;

            int bytesRead;
            while ((bytesRead = await innerStream.ReadAsync(buffer, 0, buffer.Length)) != 0)
            {
                await stream.WriteAsync(buffer, 0, bytesRead);
                uploadedBytes += bytesRead;
                if (totalLength > 0)
                {
                    this._progress((int)((uploadedBytes * 100) / totalLength));
                }
            }
        }

        protected override bool TryComputeLength(out long length)
        {
            length = this._innerContent.Headers.ContentLength ?? -1;
            return length != -1;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                this._innerContent.Dispose();
            }
            base.Dispose(disposing);
        }

    }
}