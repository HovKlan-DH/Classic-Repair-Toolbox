using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The Submit dialog: confirm what will be sent, then send it, then say what happened.
    //
    // ONE WINDOW, THREE PANELS, swapped in code rather than a wizard - the dialog stays in one
    // place on screen and Cancel keeps one meaning throughout.
    //
    // WHAT IS DECIDED HERE: nothing. The manifest is built by SubmissionManifestBuilder, the
    // hashing and HTTP are SubmissionClient's, the progress text is SubmissionProgress'. This file
    // wires controls to those and marshals back to the UI thread. That is the same rule the rest
    // of the codebase follows, and it is why the interesting parts are unit tested while this is
    // verified by running the app.
    //
    // *** NOTHING LONG-RUNNING TOUCHES THE UI THREAD. *** "Do not let submission block the UI" is
    // a named trap in the strategy document. Hashing a 76 MB board and uploading it both happen
    // on the thread pool; progress arrives through an IProgress whose callback posts to the
    // dispatcher.
    // ###########################################################################################
    public partial class SubmitDraftWindow : Window
    {
        private SubmissionIdentity? thisIdentity;
        private BoardData? thisMergedData;

        // The signed-in maintainer this submission goes with, or null for the ordinary contributor
        // (2026-10-01) - see Initialize. Set on the UI thread before the send starts; immutable.
        private ReviewSession? thisAccount;

        // The draft's KiCad calibrations, which BoardData cannot carry - see Initialize.
        private IReadOnlyList<KiCadCalibrationEntry> thisCalibrations = [];

        // What the draft held when Submit was pressed - see Initialize and NewReceipt.
        private string thisDraftFingerprint = string.Empty;

        // TWO roots, because a board's files live in two places - see SubmissionFileLocator.
        // thisBoardFolder is the DRAFT folder for this board; thisDataRoot is the synced Data/
        // tree that holds everything already published.
        private string thisBoardFolder = string.Empty;
        private string thisDataRoot = string.Empty;
        private string thisBoardDisplayName = string.Empty;

        private CancellationTokenSource? thisCancellation;

        // ###########################################################################################
        // The board's KiCad data - the "KiCad data" folder's files, which no row cites and
        // CollectReferencedFiles therefore cannot find (owner decision, 2026-09-26). The union of
        // the draft's folder and the synced official one, the draft winning, which is the same
        // draft-first rule the upload's own file resolution applies.
        // ###########################################################################################
        private IReadOnlyList<string> CollectKiCadFiles()
        {
            if (this.thisIdentity is null)
                return [];

            string officialFolder = string.IsNullOrWhiteSpace(this.thisDataRoot)
                ? string.Empty
                : Path.Combine(this.thisDataRoot, this.thisIdentity.BoardId.Replace('/', Path.DirectorySeparatorChar));

            return SubmissionKiCadFiles.Collect(this.thisIdentity.BoardId, this.thisBoardFolder, officialFolder);
        }

        // Set once the submission has been sent, so the caller can tell whether the draft should
        // be left alone (always, for now - see the Drafts tab) and what to show afterwards.
        public SubmissionResult? Result { get; private set; }

        public SubmitDraftWindow()
        {
            this.InitializeComponent();

            // Tunnel, not a plain KeyDown - a focused Button handles Enter itself and marks the
            // event handled, so a bubbling handler never runs. Same trap DiscardDraftWindow and
            // DeleteWorkbookWindow document.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            this.Closing += this.OnWindowClosing;
        }

        // ###########################################################################################
        // Prepares the dialog. Takes the merged board data rather than a draft, because what gets
        // submitted is the board as it should READ afterwards - the server diffs it against the
        // base revision itself.
        // ###########################################################################################
        public void Initialize(
            string boardDisplayName,
            BoardData mergedData,
            SubmissionIdentity identity,
            string boardFolder,
            string dataRoot,

            // The draft's KiCad calibrations. They CANNOT come from mergedData - BoardData has no
            // calibration section at all - so they travel separately. Optional and trailing, so
            // existing callers and tests are unaffected; a caller that omits them submits none,
            // which is what happened for every submission before 2026-09-22.
            IReadOnlyList<KiCadCalibrationEntry>? calibrations = null,

            // The maintainer signed in on the Maintainer tab, or null (2026-10-01) - see below.
            ReviewSession? signedIn = null,

            // What the draft holds as it is sent (DraftFingerprint, 2026-10-03), kept on the
            // receipt so the Drafts tab can grey Submit out while the draft stays the same.
            string draftFingerprint = "")
        {
            ArgumentNullException.ThrowIfNull(mergedData);
            ArgumentNullException.ThrowIfNull(identity);

            this.thisDraftFingerprint = draftFingerprint ?? string.Empty;
            this.thisBoardDisplayName = boardDisplayName;
            this.thisMergedData = mergedData;
            this.thisIdentity = identity;
            this.thisCalibrations = calibrations ?? [];
            this.thisBoardFolder = boardFolder;
            this.thisDataRoot = dataRoot;

            this.BoardNameText.Text = boardDisplayName;

            // ###########################################################################################
            // ONE EMAIL ADDRESS FOR THE WHOLE APPLICATION (owner request, 2026-09-22).
            //
            // *** THE SETTING ALREADY EXISTED AND THIS SCREEN SIMPLY DID NOT READ IT. ***
            // UserSettings.ContactEmail has been written by the Feedback tab since long before
            // submissions existed, so somebody who had already given their address there was asked
            // for it again here, with no indication the app knew it perfectly well.
            //
            // Prefilled rather than locked: the box stays editable, because a contributor may
            // genuinely want a different address on a contribution than on a bug report, and the
            // submit path is where an address is actually acted on by a person.
            //
            // Not overwritten if something is already in the box - Initialize could be called on a
            // dialog a caller has pre-populated, and silently replacing a caller's value would be
            // the kind of surprise that is very hard to see in a UI.
            //
            // *** EXCEPT BY A SIGNED-IN MAINTAINER'S ACCOUNT (owner request, 2026-10-01: "When I am
            // a maintainer, and I have logged in, then I want to use that email address everywhere
            // in the CRT app"). *** The submission then goes WITH the account (SubmitAsync sends its
            // token), so the account's address is the one the server records whatever the box says
            // - the box shows it, read only, and the note under it says why and how to change it.
            ContactAddress address = ContactAddress.Choose(signedIn, UserSettings.ContactEmail, DateTimeOffset.UtcNow);

            this.thisAccount = address.Account;

            if (address.IsFromAccount)
            {
                this.EmailTextBox.Text = address.Email;
                this.EmailTextBox.IsReadOnly = true;
                this.EmailNoteText.Text = ContactAddress.SubmitNote;
            }
            else if (string.IsNullOrWhiteSpace(this.EmailTextBox.Text))
            {
                this.EmailTextBox.Text = address.Email;
            }

            this.BuildSummary();
            this.UpdateSubmitEnabled();
        }

        // ###########################################################################################
        // The "what will be sent" panel.
        //
        // Counted from the manifest's own contents rather than from the draft, so what is shown is
        // exactly what will go - a summary computed from a different source could drift from the
        // payload and quietly become a lie.
        // ###########################################################################################
        private void BuildSummary()
        {
            this.SummaryPanel.Children.Clear();

            if (this.thisMergedData is null)
                return;

            IReadOnlyList<string> files = SubmissionManifestBuilder.CollectReferencedFiles(this.thisMergedData);

            this.AddSummaryLine("Schematics", this.thisMergedData.Schematics.Count.ToString(CultureInfo.InvariantCulture));
            this.AddSummaryLine("Components", this.thisMergedData.Components.Count.ToString(CultureInfo.InvariantCulture));
            this.AddSummaryLine("Highlights", this.thisMergedData.ComponentHighlights.Count.ToString(CultureInfo.InvariantCulture));
            this.AddSummaryLine("Files referenced", files.Count.ToString(CultureInfo.InvariantCulture));

            // The board's KiCad data travels too (2026-09-26) - only said when there is any, since
            // most boards have none and a permanent "KiCad files: 0" would read as something missing.
            int kiCadFiles = this.CollectKiCadFiles().Count;

            if (kiCadFiles > 0)
                this.AddSummaryLine("KiCad files", kiCadFiles.ToString(CultureInfo.InvariantCulture));

            // How much actually uploads is not known until the server has been asked, so this says
            // so rather than guessing - an estimate that turns out wrong is worse than no estimate.
            this.AddSummaryLine(
                "To upload",
                "worked out when you submit - files the server already has are not sent again");
        }

        private void AddSummaryLine(string label, string value)
        {
            var panel = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 6 };

            panel.Children.Add(new TextBlock
            {
                Text = $"{label}:",
                FontWeight = FontWeight.Bold,
                MinWidth = 120
            });

            panel.Children.Add(new TextBlock { Text = value, TextWrapping = TextWrapping.Wrap });

            this.SummaryPanel.Children.Add(panel);
        }

        // ###########################################################################################
        // Submit stays disabled until there is a contact address and a summary.
        //
        // Both are required by the server, so enabling the button without them would produce a
        // round trip that can only fail - and a refusal from the server reads as something being
        // wrong, rather than as a field not yet filled in.
        // ###########################################################################################
        private void OnFormChanged(object? sender, TextChangedEventArgs e)
        {
            this.UpdateSubmitEnabled();
        }

        private void UpdateSubmitEnabled()
        {
            bool hasEmail = EmailAddressRules.IsPlausible(this.EmailTextBox.Text);
            bool hasSummary = !string.IsNullOrWhiteSpace(this.SummaryTextBox.Text);

            this.SubmitButton.IsEnabled = hasEmail && hasSummary;
        }

        // ###########################################################################################
        // Does the whole submission, off the UI thread.
        // ###########################################################################################
        private async void OnSubmitClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisMergedData is null || this.thisIdentity is null)
                return;

            this.ShowProgressPanel();

            // Disposed in the finally below rather than held for the window's life: a
            // CancellationTokenSource owns a wait handle, and this dialog can be opened repeatedly.
            this.thisCancellation?.Dispose();
            this.thisCancellation = new CancellationTokenSource();

            CancellationToken token = this.thisCancellation.Token;

            var progress = new Progress<SubmissionProgress>(this.OnProgress);

            try
            {
                SubmissionResult result = await Task.Run(
                    () => this.SubmitAsync(progress, token), token);

                this.Result = result;
                this.ShowOutcome(result);
            }
            catch (OperationCanceledException)
            {
                // Cancelling is safe: an unfinalised submission is collected by the server after
                // its upload window, and the draft is untouched.
                this.Close();
            }
            catch (SubmissionRejectedException ex)
            {
                this.ShowFailure(ex.Message, ex.Findings);
            }
            catch (Exception ex) when (ex is System.Net.Http.HttpRequestException or IOException or TaskCanceledException)
            {
                this.ShowFailure(
                    "The submission could not be sent. Check your internet connection and try again - " +
                    "nothing has been lost, and your draft is unchanged.",
                    []);
            }
            finally
            {
                this.thisCancellation?.Dispose();
                this.thisCancellation = null;
            }
        }

        // ###########################################################################################
        // The three steps, in order. Runs entirely on the thread pool.
        // ###########################################################################################
        private async Task<SubmissionResult> SubmitAsync(
            IProgress<SubmissionProgress> progress,
            CancellationToken token)
        {
            // What the rows cite, plus the board's KiCad data - the folder no row cites, which the
            // rows-only list silently left behind until 2026-09-26 (a new board published without
            // its traces). Both are hashed the same way; SubmissionFileLocator resolves each path
            // against the draft first, then the synced data.
            IReadOnlyList<string> kiCadFiles = this.CollectKiCadFiles();

            List<string> paths =
            [
                .. SubmissionManifestBuilder.CollectReferencedFiles(this.thisMergedData!),
                .. kiCadFiles
            ];

            FileHashResult hashed = await SubmissionClient.HashFilesAsync(
                this.thisDataRoot, this.thisBoardFolder, paths, progress, token);

            SubmissionIdentity identity = this.thisIdentity!;

            // The summary and contact address were typed after Initialize ran, so they are read
            // from the fields captured on the UI thread rather than from the identity.
            var manifest = SubmissionManifestBuilder.Build(
                this.thisMergedData!,
                SubmitDraftWindow.IdentityToSend(identity, this.SummaryText, DateTimeOffset.UtcNow),
                hashed.Hashes,
                renames: null,
                calibrations: this.thisCalibrations,
                kiCadFiles: kiCadFiles);

            manifest.Manifest.ContactEmail = this.EmailText;

            // A file the data names but that is not on disk stops the submission here, where the
            // message can name it - rather than at finalise, where the server can only say a file
            // never arrived.
            if (!manifest.IsComplete)
                throw new SubmissionRejectedException(0, SubmitDraftWindow.AsFindings(manifest.Problems), "Some files are missing.");

            progress.Report(new SubmissionProgress(SubmissionPhase.Negotiating, 0, 0, 0, 0, null));

            var client = new SubmissionClient();

            // With a signed-in maintainer's token the submission is the account's (2026-10-01). A
            // token the server no longer accepts makes it an ordinary submission from the same
            // address - never a refusal.
            HashNegotiationResponse negotiation = await client.CreateAsync(
                manifest.Manifest, token, bearerToken: this.thisAccount?.BearerToken);

            // ###########################################################################################
            // THE RECEIPT IS WRITTEN HERE, NOT AT FINALISE.
            //
            // This is the one moment the capability token exists: the server returns it once and
            // never again, and with no account it is the ONLY thing that can later prove this
            // machine owns the submission. Recording it after the upload loop would lose it for
            // exactly the submissions a contributor most wants to look up - the ones that failed
            // partway through - and losing it is unrecoverable, since the server keeps only its
            // hash.
            //
            // Written before any byte is uploaded, so a crash mid-upload still leaves a receipt.
            // ###########################################################################################
            SubmissionReceiptStore.Record(this.NewReceipt(negotiation.SubmissionId, negotiation.UploadToken, DateTimeOffset.UtcNow));

            // Upload only what the server asked for.
            var byHash = manifest.Manifest.Files
                .GroupBy(file => file.Sha256, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            long totalBytes = negotiation.TotalBytesToUpload;
            int done = 0;

            var overall = new SubmissionProgress(
                SubmissionPhase.Uploading, 0, negotiation.MissingHashes.Count, 0, totalBytes, null);

            foreach (string hash in negotiation.MissingHashes)
            {
                token.ThrowIfCancellationRequested();

                if (!byHash.TryGetValue(hash, out SubmissionFile? file))
                    continue;

                // The same two-root resolution the hashing pass used, and its result is CHECKED.
                // An earlier version discarded it, which meant a path the rules refused was still
                // handed to the uploader as an empty string - so a broken row would have been
                // reported as a failed upload rather than as the bad path it is. Reaching here at
                // all means hashing already located this file, so a failure is a real change on
                // disk between the two passes.
                if (!SubmissionFileLocator.TryLocate(
                        this.thisDataRoot, this.thisBoardFolder, file.Path,
                        out string absolute, out string reason))
                {
                    throw new SubmissionRejectedException(
                        0,
                        SubmitDraftWindow.AsFindings([$"[{file.Path}] cannot be uploaded: {reason}"]),
                        "A file could not be read.");
                }

                overall = overall with { FilesDone = done, CurrentFile = Path.GetFileName(absolute) };
                progress.Report(overall);

                await client.UploadBlobAsync(
                    negotiation.SubmissionId, negotiation.UploadToken, hash, absolute,
                    progress, overall, token);

                // The bytes as well as the count - without them every file began again from
                // "0 bytes" and the bar never moved past the file in hand.
                overall = overall.AfterFileSent(file.SizeBytes);
                done++;
            }

            progress.Report(new SubmissionProgress(SubmissionPhase.Finalising, done, done, totalBytes, totalBytes, null));

            SubmissionResult result = await client.FinaliseAsync(negotiation.SubmissionId, negotiation.UploadToken, token);

            // The receipt learns the state the server just confirmed, so the Drafts tab's badge says
            // it at once instead of "Not checked yet" until the next launch.
            SubmissionReceiptStore.RecordFinalised(result, DateTimeOffset.UtcNow);

            return result;
        }

        // ###########################################################################################
        // The receipt kept for a submission the server has just created - see SubmitAsync for why
        // it is written then. It carries what the draft held as it was sent (2026-10-03).
        // ###########################################################################################
        internal SubmissionReceipt NewReceipt(long submissionId, string uploadToken, DateTimeOffset sentUtc) => new()
        {
            SubmissionId = submissionId,
            UploadToken = uploadToken,
            BoardId = this.thisIdentity?.BoardId ?? string.Empty,
            Summary = this.SummaryText,
            SentUtc = sentUtc,
            DraftFingerprint = this.thisDraftFingerprint
        };

        // Read on the UI thread before the work starts, because a TextBox cannot be touched from
        // the thread pool.
        private string SummaryText => this.thisSummaryText;

        private string EmailText => this.thisEmailText;

        private string thisSummaryText = string.Empty;
        private string thisEmailText = string.Empty;

        // ###########################################################################################
        // Drives the "read the boxes and remember the address" half of ShowProgressPanel from a
        // headless test, without the panel switching a test cannot meaningfully assert on.
        //
        // This is the SHIPPED path, not a parallel one: ShowProgressPanel calls exactly this. The
        // split exists because the rest of that method hides and shows panels on a window that is
        // never shown, which would prove nothing.
        // ###########################################################################################
        internal void CaptureEntriesForTests() => this.CaptureEntries();

        private void ShowProgressPanel()
        {
            this.CaptureEntries();

            this.ConfirmPanel.IsVisible = false;
            this.ProgressPanel.IsVisible = true;
            this.SubmitButton.IsVisible = false;

            // "Preparing" until the upload starts - OnProgress moves it on with the phase.
            this.HeaderText.Text = SubmissionProgress.Starting().Heading;
        }

        private void CaptureEntries()
        {
            this.thisSummaryText = this.SummaryTextBox.Text?.Trim() ?? string.Empty;
            this.thisEmailText = this.EmailTextBox.Text?.Trim() ?? string.Empty;

            // *** REMEMBERED APP-WIDE, so this is asked for once and never again. *** The same
            // UserSettings.ContactEmail the Feedback tab reads and writes, which is the whole point
            // of sharing it: an address typed here fills the Feedback tab too, and vice versa.
            //
            // Written on SEND rather than on every keystroke, so a half-typed address abandoned by
            // closing the dialog is never persisted. Only a plausible one is kept - saving
            // "dennis@" would quietly poison the prefill on every other screen.
            //
            // Never a signed-in maintainer's account address (2026-10-01): that one is the
            // account's, and signing out must bring back what the user typed (ContactAddress).
            if (this.thisAccount is null && EmailAddressRules.IsPlausible(this.thisEmailText))
            {
                UserSettings.ContactEmail = this.thisEmailText;
            }
        }

        // ###########################################################################################
        // Progress arrives from the thread pool, so every update is posted to the dispatcher.
        // ###########################################################################################
        private void OnProgress(SubmissionProgress progress)
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.ProgressText.Text = progress.Describe();

                // Only while the progress panel is up: a report still queued when the outcome
                // arrives must not turn "Contribution sent" back into "Sending contribution".
                if (this.ProgressPanel.IsVisible)
                    this.HeaderText.Text = progress.Heading;

                double? fraction = progress.Fraction;

                // Null means "nothing measurable yet" - an indeterminate bar rather than one
                // sitting at zero, which reads as stuck.
                this.ProgressBar.IsIndeterminate = fraction is null;

                if (fraction is not null)
                    this.ProgressBar.Value = fraction.Value;
            });
        }

        private void ShowOutcome(SubmissionResult result)
        {
            // ###########################################################################################
            // LOGGED, because a refusal the user cannot act on is the worst outcome this dialog has.
            //
            // The first real submission came back "not accepted" with an EMPTY findings panel - the
            // server had recorded three findings and the contributor was shown none of them. There
            // was no way to tell from the client whether the list arrived empty or failed to
            // render, because nothing recorded what actually came back.
            // ###########################################################################################
            Logger.Info(
                $"Submission [{result.SubmissionId}] outcome: accepted [{result.IsAccepted}] " +
                $"state [{result.State}] findings [{result.Findings?.Count ?? 0}]");

            foreach (ValidationFinding finding in result.Findings ?? [])
                Logger.Info($"  finding [{finding.Code}] [{finding.Subject}] {finding.Message}");

            this.ProgressPanel.IsVisible = false;
            this.OutcomePanel.IsVisible = true;
            this.CancelButton.Content = "Close";

            if (result.IsAccepted)
            {
                this.HeaderText.Text = "Contribution sent";

                this.OutcomeText.Text =
                    "Your contribution is queued for review. You will get an email when it has been " +
                    "looked at.\n\nYour draft is untouched - keep using it as normal. It stays until " +
                    "your work has been published.";
            }
            else
            {
                this.HeaderText.Text = "Contribution not accepted";

                this.OutcomeText.Text =
                    "The server found problems that have to be fixed before this can be reviewed. " +
                    "Your draft is untouched.";
            }

            this.ShowFindings(result.Findings, isRefusal: !result.IsAccepted);
        }

        private void ShowFailure(string message, IReadOnlyList<ValidationFinding> findings)
        {
            this.ProgressPanel.IsVisible = false;
            this.OutcomePanel.IsVisible = true;
            this.CancelButton.Content = "Close";
            this.HeaderText.Text = "Contribution not sent";
            this.OutcomeText.Text = message;

            this.ShowFindings(findings, isRefusal: true);
        }

        // ###########################################################################################
        // Shows what the server (or the local check) found wrong.
        //
        // EVERY finding, each naming its subject - a contributor with 400 components cannot act on
        // "something is invalid". Errors first, because those are what blocked it.
        // ###########################################################################################
        // Drives the real findings list from a headless test. The dialog's own path into it goes
        // through an HTTP round trip, which a test must not make - but WHAT IS RENDERED is entirely
        // this method's decision, and it is where a white-on-white bug hid in plain sight.
        internal void ShowFindingsForTests(IReadOnlyList<ValidationFinding>? findings, bool isRefusal = true)
        {
            this.ShowFindings(findings, isRefusal);
        }

        private void ShowFindings(IReadOnlyList<ValidationFinding>? findings, bool isRefusal)
        {
            this.FindingsPanel.Children.Clear();

            // ###########################################################################################
            // "PROBLEMS WERE FOUND" WITH NOTHING LISTED IS A DEAD END, and it shipped that way: the
            // first real submission showed exactly that, while the server had recorded three
            // findings naming precisely what was wrong.
            //
            // A rejection the contributor cannot act on is worse than a crash - a crash at least
            // looks like a fault, whereas this looks like their data is wrong and gives them
            // nothing to fix. So when a refusal carries no findings, SAY that rather than showing
            // an empty space, and point at the one thing they can actually do.
            // ###########################################################################################
            // ###########################################################################################
            // THE APOLOGY IS FOR A REFUSAL ONLY. An ACCEPTED submission legitimately carries no
            // findings - that is what "nothing wrong with it" looks like - and this panel is shared
            // by both outcomes.
            //
            // Called unconditionally, "the server did not say what was wrong" appeared underneath
            // "Contribution sent", telling a contributor whose work had just been accepted that
            // something had gone wrong. Reported on the first successful submission, which is the
            // worst possible moment to undermine.
            //
            // Note this returns rather than skipping the whole method: an accepted submission can
            // still carry WARNINGS (a missing summary, say), and those are worth showing. Only the
            // "something went wrong" message is refusal-only.
            // ###########################################################################################
            if (findings is null || findings.Count == 0)
            {
                if (!isRefusal)
                    return;

                this.FindingsPanel.Children.Add(new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    Text = "The server did not say what was wrong, which is a fault in CRT rather " +
                           "than in your data. Please report this - the details are in the log."
                });

                return;
            }

            foreach (ValidationFinding finding in findings
                .OrderByDescending(finding => finding.Severity == ValidationSeverity.Error))
            {
                var text = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    Text = string.IsNullOrWhiteSpace(finding.Subject)
                        ? finding.Message
                        : $"{finding.Subject}: {finding.Message}"
                };

                // ###########################################################################################
                // *** NOT Button_Cancel_Fg. *** That key is WHITE - it is the foreground for text on
                // a red-filled Cancel BUTTON, not for text on this panel's ordinary background.
                //
                // Using it here rendered every finding in white on white. The three findings from
                // the first real submission were all present, laid out and completely invisible,
                // which read as "the server found problems" followed by an empty box - the single
                // most useless thing this dialog could do.
                //
                // Button_Cancel_Bg is the RED, and it is legible on this background, so that is the
                // one to colour text with. Falling back to leaving Foreground ALONE (inheriting the
                // window's own) rather than assigning null: a null Foreground is not "default", and
                // it is how the text vanished in the first place.
                // ###########################################################################################
                if (finding.Severity == ValidationSeverity.Error
                    && this.TryFindResource("Button_Cancel_Bg", out object? brush)
                    && brush is IBrush errorBrush)
                {
                    text.Foreground = errorBrush;
                }

                this.FindingsPanel.Children.Add(text);
            }
        }

        // ###########################################################################################
        // The identity the manifest is built from: the one Initialize was given, with the summary
        // typed since and the moment of sending. Everything else is COPIED - a field left out here
        // never leaves this computer, however the caller filled it in (a new board's notes nearly
        // did, 2026-10-05), so SubmitDraftWindowIdentityTests checks every property by reflection.
        // ###########################################################################################
        internal static SubmissionIdentity IdentityToSend(SubmissionIdentity identity, string summary, DateTimeOffset nowUtc)
        {
            ArgumentNullException.ThrowIfNull(identity);

            return new SubmissionIdentity
            {
                BoardId = identity.BoardId,
                Manufacturer = identity.Manufacturer,
                Hardware = identity.Hardware,
                Board = identity.Board,
                BaseRevision = identity.BaseRevision,
                Summary = summary,
                HardwareNotes = identity.HardwareNotes,
                ApplicationVersion = identity.ApplicationVersion,
                CreatedUtc = nowUtc
            };
        }

        private static IReadOnlyList<ValidationFinding> AsFindings(IReadOnlyList<string> problems)
        {
            return problems
                .Select(problem => new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = "local.file_missing",
                    Subject = string.Empty,
                    Message = problem
                })
                .ToList();
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                this.thisCancellation?.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // The submission already finished; there is nothing to cancel.
            }

            this.Close();
        }

        // Escape cancels. Enter does NOT submit: this dialog sends work to a public server, and a
        // reflexive Enter while filling in the summary box should not do that.
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
        }

        // Closing by the title bar does not go through Cancel, so the upload is stopped here too.
        private void OnWindowClosing(object? sender, WindowClosingEventArgs e)
        {
            // May already be disposed if the submission finished - cancelling a disposed source
            // throws, and a throw from a Closing handler would leave the window stuck open.
            try
            {
                this.thisCancellation?.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }
        }
    }
}
