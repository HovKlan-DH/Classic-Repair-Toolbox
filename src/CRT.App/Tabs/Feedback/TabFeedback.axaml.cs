using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.OnlineHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    public partial class TabFeedback : UserControl
    {
        private readonly ObservableCollection<string> _customAttachments = new();

        // Where feedback is posted: CRT.Server's route (FeedbackContract).
        internal const string FeedbackUrl = AppConfig.CrtServerBaseUrl + "/" + FeedbackContract.PathUnderApi;

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
        // *** A SIGNED-IN MAINTAINER'S ADDRESS (owner request, 2026-10-01: "When I am a maintainer,
        // and I have logged in, then I want to use that email address everywhere in the CRT app -
        // e.g. for the Feedback tab"). *** Main hands the sign-in over (ShareMaintainerSignIn); while
        // there is one the box shows the account's address, read only, with a note saying why. It
        // is never saved as the typed address, so signing out brings back what was typed before
        // (ContactAddress has the rule).
        // ###########################################################################################
        private ReviewSession? thisMaintainerAccount;
        private bool thisShowsAccountAddress;

        internal void UseMaintainerAccount(ReviewSession? account)
        {
            if (Equals(this.thisMaintainerAccount, account))
                return;

            this.thisMaintainerAccount = account;

            ContactAddress address = ContactAddress.Choose(account, UserSettings.ContactEmail, DateTimeOffset.UtcNow);

            // Signed out again: what was typed comes back. Signed in: the account's address.
            if (address.IsFromAccount || this.thisShowsAccountAddress)
                this.EmailTextBox.Text = address.Email;

            this.thisShowsAccountAddress = address.IsFromAccount;
            this.EmailTextBox.IsReadOnly = address.IsFromAccount;
            this.EmailAccountNoteText.Text = address.IsFromAccount ? ContactAddress.FeedbackNote : string.Empty;
            this.EmailAccountNoteText.IsVisible = address.IsFromAccount;
        }

        // ###########################################################################################
        // Persists the shared email address when the field loses focus and the value is valid.
        // ###########################################################################################
        private void OnEmailTextBoxLostFocus(object? sender, RoutedEventArgs e)
        {
            // The account's address is the account's, never the typed one.
            if (this.thisShowsAccountAddress)
                return;

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

            if (!this.thisShowsAccountAddress)
                UserSettings.ContactEmail = email;

            if (string.IsNullOrEmpty(feedback))
            {
                this.ShowStatus("Please provide a description of your issue or suggestion before sending", isError: true);
                return;
            }

            this.SubmitButton.IsEnabled = false;
            this.ShowStatus(string.Empty, isError: false);

            bool attachLogs = this.AttachLogfileCheckBox.IsChecked == true;
            bool attachConfig = this.AttachConfigsCheckBox.IsChecked == true;
            var customPaths = this._customAttachments.ToList();

            try
            {
                // ###########################################################################################
                // *** UNDER THE "PLEASE WAIT" OVERLAY (owner request, 2026-09-28: "this should also be
                // visible then when submitting a large feedback in the Feedback tab"). *** The packing
                // and upload percentages go onto the overlay, and each one starts its two minutes
                // again - so a large attachment on a slow line is never cut off while it is visibly
                // moving, and only two minutes with NOTHING happening gives up (WaitLimit).
                // ###########################################################################################
                WaitResult<(bool Success, int StatusCode, string ResponseBody)> waited = await BusyOverlay.RunAsync(
                    this,
                    CrtWaitWording.SendingFeedback,
                    context =>
                    {
                        IProgress<string> progress = context.AsProgress();
                        return Task.Run(() => this.ProcessAndSendFeedbackAsync(
                            email, feedback, attachLogs, attachConfig, customPaths, progress, context.Token));
                    });

                // Nothing answers back from the feedback page, so whether it arrived cannot be
                // checked - it MAY have. The text stays, so nothing typed is lost either way.
                if (waited.IsTimedOut)
                {
                    Logger.Warning("Feedback submission: no answer within the wait limit.");
                    this.ShowStatus(CrtWaitWording.FeedbackNoAnswer, isError: true);
                    return;
                }

                var (success, statusCode, responseBody) = waited.Value;

                if (success)
                {
                    this.ShowStatus(FeedbackWording.Sent, isError: false);
                    this.FeedbackTextBox.Text = string.Empty;
                    this._customAttachments.Clear();
                    this.AttachLogfileCheckBox.IsChecked = false;
                    this.AttachConfigsCheckBox.IsChecked = false;
                }
                else
                {
                    Logger.Warning($"Feedback submission failed. HTTP {statusCode}. Server responded with: {responseBody}");
                    this.ShowStatus(FeedbackWording.Failed(statusCode), isError: true);
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
        private async Task<(bool Success, int StatusCode, string ResponseBody)> ProcessAndSendFeedbackAsync(string email, string feedbackText, bool attachLogs, bool attachConfig, List<string> customPaths, IProgress<string> progress, CancellationToken cancellationToken)
        {
            var targetFiles = new List<(string Source, string ZipEntryName)>();
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string localAppFolder = Path.Combine(appData, AppConfig.AppFolderName);

            // 1. Gather all files to be zipped
            if (attachLogs)
            {
                targetFiles.Add((Path.Combine(localAppFolder, AppConfig.LogFileName), AppConfig.LogFileName));

                // The crash file rides along with the log, because it is the ONLY record that
                // survives a restart - the log is truncated on every launch, and a user reporting a
                // crash has almost always relaunched before they get here, so the log they attach
                // frequently no longer contains the crash they are writing about.
                //
                // Added unconditionally: it does not exist on an installation that has never
                // crashed, and AddFileToZipSafe already skips a missing source rather than failing
                // the whole submission.
                targetFiles.Add((Path.Combine(localAppFolder, AppConfig.CrashFileName), AppConfig.CrashFileName));
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

            // 3. Zip the files into a TEMPORARY FILE, not into memory (2026-10-03): the attachments
            //    may be 250 MB packed ("one could potentially zip the entire board"), and memory
            //    held up to three copies of them before. DeleteOnClose removes it however this ends.
            string zipPath = Path.Combine(Path.GetTempPath(), $"CRT-feedback-{Guid.NewGuid():N}.zip");
            await using var zipFile = new FileStream(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.DeleteOnClose);
            using (var archive = new ZipArchive(zipFile, ZipArchiveMode.Create, leaveOpen: true))
            {
                long currentUncompressedBytes = 0;
                int lastReportedPercent = -1;

                if (totalUncompressedBytes == 0)
                {
                    progress.Report("Packaging payload... 100%");
                }

                foreach (var file in targetFiles)
                {
                    AddFileToZipSafe(archive, file.Source, file.ZipEntryName, bytesAdded =>
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

            zipFile.Position = 0;

            // An empty archive is just its 22-byte end record - nothing attached.
            bool hasAttachment = zipFile.Length > 22;

            // Over the server's limit: said here, before a long upload ends in a refusal.
            if (hasAttachment && zipFile.Length > FeedbackContract.MaximumAttachmentBytes)
                return (false, 413, $"Not sent: the zip is {zipFile.Length} bytes.");

            // 4. The form - CRT.Data's FeedbackContract, the one the server reads. The zip is
            //    streamed from its file as it is sent.
            using var httpClient = new HttpClient { Timeout = AppConfig.UploadTimeout };
            using MultipartFormDataContent formContent = FeedbackContract.BuildForm(
                email, feedbackText, AppConfig.AppDisplayVersionString, hasAttachment ? zipFile : null);

            // Track Upload progress - reported often enough to keep the wait alive on a slow line,
            // and saying what is waited for once the last byte is sent (ProgressableStreamContent).
            using var progressContent = new ProgressableStreamContent(formContent, percent => progress.Report(CrtWaitWording.SendingFeedbackAt(percent)));

            // ###########################################################################################
            // *** CRT.SERVER SINCE 2026-10-03, NOT THE OLD FEEDBACK ADDRESS. *** The same form and
            // the same "Success" answer; older CRTs still post to the old address, which Apache
            // forwards to the same route.
            // ###########################################################################################
            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("CRT "+ AppConfig.AppDisplayVersionString);
            var response = await httpClient.PostAsync(TabFeedback.FeedbackUrl, progressContent, cancellationToken);

            string responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

            return (FeedbackContract.IsSuccess((int)response.StatusCode, responseBody), (int)response.StatusCode, responseBody);
        }

        // ###########################################################################################
        // Adds a local file path to the provided ZipArchive, silencing read-access violations 
        // if file doesn't exist or is locked. Reports bytes copied back out to optionally track 
        // compression progress dynamically.
        // ###########################################################################################
        private static void AddFileToZipSafe(ZipArchive archive, string sourcePath, string entryName, Action<int>? onBytesRead = null)
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
}
