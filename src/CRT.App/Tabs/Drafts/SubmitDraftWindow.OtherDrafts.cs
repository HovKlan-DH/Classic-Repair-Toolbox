using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE SUBMIT DIALOG'S OTHER DRAFTS (CRT 4.0 step 2, slice 5 - the plan: "Submit offers to send all
    // the drafts at once, with one message").
    //
    // The contributor's other drafts are listed under the message, UNTICKED - sending is theirs to
    // choose - each greyed out with the reason when its row says it cannot go. Every ticked one is
    // sent AFTER this board's, as a submission of its OWN (the server takes one board at a time, and a
    // maintainer reviews one board), with the same email address and the same message. The Contribute
    // tab makes each ready only once Submit is pressed (`thisPrepareOther`: its draft saved, looked at
    // again, credited the same way while "Credit me on the other drafts" is ticked) - a draft that
    // turns out unable to go is said, never stops the rest. What came of each board is listed by name
    // when several went, and credits put into drafts before a send that failed are said.
    //
    // Nothing here decides what a submission holds: each is sent by the same SubmitAsync as one board.
    // ###########################################################################################
    public partial class SubmitDraftWindow
    {
        // How the Contribute tab makes another draft ready to send, with the credit name and contact
        // typed here - see OfferOtherDrafts.
        private Func<HardwareBoardEntry, string, string, Task<OtherDraftReady>>? thisPrepareOther;

        // The board being sent now when several go, to name it in the progress line - read and set on
        // the UI thread only. Empty for one board, whose progress line names none.
        private string thisSendingName = string.Empty;

        // How many credits this board's own work deserves (ShowCreditOffer) - the credit offer shows
        // while there are any, or while another draft is ticked.
        private int thisOwnCreditSuggestions;

        // Stands in for sending one board (SubmitAsync, which goes to the server) - a test sends
        // nothing over the network (test rule 6). null is the shipped path.
        internal Func<SubmitDraftTarget, CancellationToken, Task<SubmissionResult>>? SendOverrideForTests { get; set; }

        // ###########################################################################################
        // The other drafts on offer, each a tick box - unticked; one that cannot go greyed out with
        // its reason. `prepare` makes a ticked one ready once Submit is pressed. Nothing to offer
        // hides the list.
        // ###########################################################################################
        internal void OfferOtherDrafts(
            IReadOnlyList<OtherDraftOffer> offers,
            Func<HardwareBoardEntry, string, string, Task<OtherDraftReady>> prepare)
        {
            ArgumentNullException.ThrowIfNull(offers);
            ArgumentNullException.ThrowIfNull(prepare);

            this.thisPrepareOther = prepare;
            this.OtherDraftsPanel.Children.Clear();

            foreach (OtherDraftOffer offer in offers)
            {
                var box = new CheckBox
                {
                    Content = SubmitDraftsWording.Choice(offer.DisplayName, offer.RowSummary, offer.NotOffered),
                    IsChecked = false,
                    IsEnabled = offer.NotOffered is null,
                    Tag = offer,
                };

                box.IsCheckedChanged += (_, _) => this.ShowCreditOfferFor();
                this.OtherDraftsPanel.Children.Add(box);
            }

            this.OtherDraftsHeadingText.Text = SubmitDraftsWording.OtherDraftsHeading;
            this.OtherDraftsIntroText.Text = SubmitDraftsWording.OtherDraftsIntro;
            this.OtherDraftsSection.IsVisible = offers.Count > 0;
            this.ShowCreditOfferFor();
        }

        // The other drafts ticked, in the order listed.
        private List<OtherDraftOffer> TickedOtherDrafts() =>
            this.OtherDraftsPanel.Children
                .OfType<CheckBox>()
                .Where(box => box.IsChecked == true && box.IsEnabled)
                .Select(box => box.Tag)
                .OfType<OtherDraftOffer>()
                .ToList();

        // What the list shows, and whether each can be ticked - for tests.
        internal IReadOnlyList<(string Text, bool CanTick)> OtherDraftChoicesForTests =>
            this.OtherDraftsSection.IsVisible
                ? this.OtherDraftsPanel.Children.OfType<CheckBox>().Select(box => (box.Content as string ?? string.Empty, box.IsEnabled)).ToList()
                : [];

        internal void TickOtherDraftForTests(int index, bool ticked = true) =>
            this.OtherDraftsPanel.Children.OfType<CheckBox>().ElementAt(index).IsChecked = ticked;

        // ###########################################################################################
        // Every credit of this board unticked: no credit is wanted, so the other drafts' is unticked
        // too (code review, 2026-10-26 - the other drafts were credited under the name all the same).
        // Ticked again by hand, it stays.
        // ###########################################################################################
        private void OnOwnCreditTicked()
        {
            List<CheckBox> own = this.CreditChoicesPanel.Children.OfType<CheckBox>().ToList();

            if (own.Count > 0 && own.All(box => box.IsChecked != true))
            {
                this.CreditOthersCheckBox.IsChecked = false;
            }
        }

        // The credit offer: this board's suggestions, and - while another draft is ticked - whether the
        // other drafts are credited under the name too (a tick box of its own).
        private void ShowCreditOfferFor()
        {
            bool othersTicked = this.TickedOtherDrafts().Count > 0;

            this.CreditChoicesPanel.IsVisible = this.thisOwnCreditSuggestions > 0;
            this.CreditOthersText.Text = SubmitDraftsWording.CreditOthers;
            this.CreditOthersCheckBox.IsVisible = othersTicked;
            this.CreditOfferPanel.IsVisible = this.thisOwnCreditSuggestions > 0 || othersTicked;
        }

        // ###########################################################################################
        // What came of sending one board. A submission the server answered carries its result (taken or
        // turned down); one that never got there carries why - in a sentence for one board alone
        // (ShowFailure, as ever), and in the short form "Board: not sent - <reason>" when several went.
        // ###########################################################################################
        private sealed record BoardOutcome(
            string BoardName,
            SubmissionResult? Result,
            string FailureMessage,
            string NotSentReason,
            IReadOnlyList<ValidationFinding> Findings)
        {
            public static BoardOutcome NotSent(string boardName, string reason) =>
                new(boardName, null, reason, reason, []);
        }

        // ###########################################################################################
        // Sends one board and says what came of it. Only the contributor's cancel leaves here (code
        // review, 2026-10-13): every other way it can end is an outcome, so one board that cannot go
        // never stops the next.
        // ###########################################################################################
        private async Task<BoardOutcome> SendOneAsync(SubmitDraftTarget target, IProgress<SubmissionProgress> progress, CancellationToken token)
        {
            try
            {
                SubmissionResult result = this.SendOverrideForTests is { } send
                    ? await send(target, token)
                    : await Task.Run(() => this.SubmitAsync(target, progress, token), token);

                return new BoardOutcome(target.DisplayName, result, string.Empty, string.Empty, result.Findings ?? []);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                throw;
            }
            catch (SubmissionRejectedException ex)
            {
                return new BoardOutcome(target.DisplayName, null, ex.Message, SubmitDraftsWording.AsReason(ex.Message), ex.Findings);
            }
            catch (Exception ex) when (ex is System.Net.Http.HttpRequestException or IOException or TaskCanceledException)
            {
                // A request that timed out is a TaskCanceledException too - not the contributor's
                // cancel, which the first catch took - so it reads as "could not be sent".
                return new BoardOutcome(target.DisplayName, null, SubmitDraftWindow.NoConnectionMessage(this.thisCreditAdded), SubmitDraftsWording.NoConnection, []);
            }
            catch (Exception ex)
            {
                // A file the hashing cannot read (UnauthorizedAccessException is not an IOException)
                // was the known way here; anything else is reported the same way and logged in full.
                Logger.Critical($"Submitting the draft of [{target.DisplayName}] failed unexpectedly - [{ex}]");
                return new BoardOutcome(target.DisplayName, null, SubmitDraftWindow.UnexpectedMessage(this.thisCreditAdded), SubmitDraftsWording.Unexpected, []);
            }
        }

        // ###########################################################################################
        // The other drafts ticked, each made ready by the Contribute tab and sent after this board -
        // what came of each added to `outcomes`. True when any was credited under the name typed.
        // ###########################################################################################
        private async Task<bool> SendOtherDraftsAsync(
            IReadOnlyList<OtherDraftOffer> others,
            string creditName,
            string creditContact,
            List<BoardOutcome> outcomes,
            IProgress<SubmissionProgress> progress,
            CancellationToken token)
        {
            bool credited = false;

            foreach (OtherDraftOffer offer in others)
            {
                token.ThrowIfCancellationRequested();

                this.thisSendingName = string.Empty;
                this.ProgressText.Text = SubmitDraftsWording.GettingReady(offer.DisplayName);
                this.ProgressBar.IsIndeterminate = true;

                OtherDraftReady ready;

                try
                {
                    ready = this.thisPrepareOther is { } prepare
                        ? await prepare(offer.Entry, creditName, creditContact)
                        : OtherDraftReady.NotSent(SubmitDraftsWording.DraftGone);
                }
                catch (Exception ex)
                {
                    Logger.Critical($"Getting the draft of [{offer.DisplayName}] ready to send failed unexpectedly - [{ex}]");
                    ready = OtherDraftReady.NotSent(SubmitDraftsWording.Unexpected);
                }

                if (ready.Target is not { } target)
                {
                    outcomes.Add(BoardOutcome.NotSent(offer.DisplayName, ready.Reason));
                    continue;
                }

                credited |= target.Credited;
                this.thisSendingName = target.DisplayName;
                outcomes.Add(await this.SendOneAsync(target, progress, token));
            }

            return credited;
        }

        // ###########################################################################################
        // What came of several boards: how many went, then each board by name - queued, turned down
        // or not sent and why - with what the server found under it.
        // ###########################################################################################
        private void ShowOutcomes(IReadOnlyList<BoardOutcome> outcomes, bool credited)
        {
            foreach (BoardOutcome outcome in outcomes)
            {
                Logger.Info(
                    $"Submission of [{outcome.BoardName}]: sent [{outcome.Result is not null}] accepted [{outcome.Result?.IsAccepted}] " +
                    $"state [{outcome.Result?.State}] findings [{outcome.Findings.Count}] reason [{outcome.NotSentReason}]");
            }

            int sent = outcomes.Count(outcome => outcome.Result is { IsAccepted: true });

            this.ProgressPanel.IsVisible = false;
            this.OutcomePanel.IsVisible = true;
            this.CancelButton.Content = "Close";
            this.HeaderText.Text = SubmitDraftsWording.Heading(sent, outcomes.Count);
            this.OutcomeText.Text = SubmitDraftsWording.Intro(sent, credited);
            this.FindingsPanel.Children.Clear();

            foreach (BoardOutcome outcome in outcomes)
            {
                string line = outcome.Result switch
                {
                    { IsAccepted: true } => SubmitDraftsWording.Queued(outcome.BoardName),
                    not null => SubmitDraftsWording.NotAccepted(outcome.BoardName),
                    null => SubmitDraftsWording.NotSent(outcome.BoardName, outcome.NotSentReason),
                };

                this.FindingsPanel.Children.Add(new TextBlock
                {
                    Text = line,
                    TextWrapping = TextWrapping.Wrap,
                    FontWeight = FontWeight.SemiBold,
                    Margin = new Avalonia.Thickness(0, this.FindingsPanel.Children.Count == 0 ? 0 : 6, 0, 0),
                });

                foreach (ValidationFinding finding in outcome.Findings.OrderByDescending(finding => finding.Severity == ValidationSeverity.Error))
                {
                    var text = new TextBlock
                    {
                        Text = string.IsNullOrWhiteSpace(finding.Subject) ? finding.Message : $"{finding.Subject}: {finding.Message}",
                        TextWrapping = TextWrapping.Wrap,
                        Margin = new Avalonia.Thickness(12, 0, 0, 0),
                    };

                    // Button_Cancel_Bg, never Button_Cancel_Fg - see ShowFindings.
                    if (finding.Severity == ValidationSeverity.Error &&
                        this.TryFindResource("Button_Cancel_Bg", out object? brush) && brush is IBrush errorBrush)
                    {
                        text.Foreground = errorBrush;
                    }

                    this.FindingsPanel.Children.Add(text);
                }
            }
        }

        // What the outcome lists, line by line - for tests.
        internal IReadOnlyList<string> OutcomeLinesForTests =>
            this.OutcomePanel.IsVisible
                ? [this.HeaderText.Text ?? string.Empty, .. this.FindingsPanel.Children.OfType<TextBlock>().Select(text => text.Text ?? string.Empty)]
                : [];
    }

    // ###########################################################################################
    // One board to send: what the dialog needs of it - the board as it should read, who it is, where
    // its draft's files lie, its KiCad calibrations, and what the draft held as it was read (kept on
    // the receipt). `Credited`: a credit went into the draft under the name typed.
    // ###########################################################################################
    internal sealed record SubmitDraftTarget(
        string DisplayName,
        BoardData Board,
        SubmissionIdentity Identity,
        string BoardFolder,
        IReadOnlyList<KiCadCalibrationEntry> Calibrations,
        string Fingerprint,
        bool Credited = false);

    // Another draft the dialog offers: its board, its row's summary, and why it cannot go - or null.
    internal sealed record OtherDraftOffer(HardwareBoardEntry Entry, string DisplayName, string RowSummary, string? NotOffered);

    // Another draft made ready to send - or why it cannot go after all.
    internal sealed record OtherDraftReady(SubmitDraftTarget? Target, string Reason)
    {
        public static OtherDraftReady NotSent(string reason) => new(null, reason);

        public static OtherDraftReady Ready(SubmitDraftTarget target) => new(target, string.Empty);
    }
}
