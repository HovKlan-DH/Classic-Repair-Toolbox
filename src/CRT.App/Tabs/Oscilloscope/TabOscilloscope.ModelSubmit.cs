using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "SUBMIT THIS MODEL" - A MODEL OF THE USER'S OWN SENT IN FOR EVERYBODY (CRT 4.0 step 5 - the
    // plan: "Submit this model sends it to the administrator's queue on the Maintainer tab, which
    // publishes it into the main file for everybody"; the server's OscilloscopeModelFlows).
    //
    // A panel under the editor's buttons: the address the outcome is mailed to - the Maintainer tab's
    // account's while signed in there, read only, the typed one otherwise (ContactAddress, the Submit
    // dialog's and the Feedback tab's one rule) - a name if the user likes, and a message to the
    // administrator. It says whether "Run full test suite" went through on the model AS IT IS NOW
    // (a change since makes it not so); the server is told the same, decides nothing by it.
    //
    // Sent under the window's BusyOverlay; what the model's own rules or the server refuse is said in
    // the panel, which stays. Sent, the panel goes and the editor says when (UserSettings, so it is
    // still said after a restart); the outcome comes by mail. Sending again replaces the one still
    // waiting (the server's rule), so a corrected model just goes again.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        internal const string SubmitIntro =
            "The administrator looks at it and publishes it into CRT's data, so everybody with this scope can use it. You get " +
            "a mail when it is published, or why not. Your model stays here either way.";

        internal const string SuitePassedLine =
            "Run full test suite went through on this model as it is now - the administrator sees that.";

        internal const string SuiteNotRunLine =
            "Run full test suite has not gone through on this model as it is now. Run it first if you can - the administrator " +
            "sees whether it did.";

        internal const string NoAddress = "Type the email address you want to be told at.";

        internal const string BadAddress = "That does not look like an email address.";

        internal const string SendingModel = "Sending your oscilloscope model...";

        internal const string SendNoAnswer =
            "The server did not answer in time. It may have taken your model - if you hear nothing by mail, send it again.";

        // The Maintainer tab's sign-in, handed over by Main (ShareMaintainerSignIn) - the account
        // a model is sent with, and its address.
        private ReviewSession? thisMaintainerAccount;

        // The model "Run full test suite" last went through on, as its values (OscilloscopeModelValues.Fingerprint) -
        // set off the UI thread by the suite, read on it; a string swapped whole.
        private volatile string? thisSuitePassedFingerprint;

        // Answers a send without a server - for tests; sees the request and the token.
        internal Func<OscilloscopeModelSubmissionRequest, string?, Task<OscilloscopeModelSendOutcome>>? SendOverrideForTests { get; set; }

        internal void UseMaintainerAccount(ReviewSession? account)
        {
            this.thisMaintainerAccount = account;

            if (this.SubmitModelPanel.IsVisible)
            {
                this.FillSubmitAddress();
            }
        }

        // ###########################################################################################
        // The suite went through on `model` (TabOscilloscope.Capture.cs, at the end of a full run) -
        // said at once in the send panel when it is open (code review, 2026-10-29: it kept "has not
        // gone through" in the warning colour). The run ends on whichever thread it ended on.
        // ###########################################################################################
        private void NoteSuitePassed(OscilloscopeEntry model)
        {
            this.thisSuitePassedFingerprint = model.Row is MainDocument.OscilloscopeRow row ? OscilloscopeModelValues.Fingerprint(row) : null;

            if (Dispatcher.UIThread.CheckAccess())
            {
                this.ShowSubmitSuiteLineIfOpen();
            }
            else
            {
                Dispatcher.UIThread.Post(this.ShowSubmitSuiteLineIfOpen);
            }
        }

        private void ShowSubmitSuiteLineIfOpen()
        {
            if (this.SubmitModelPanel.IsVisible)
            {
                this.ShowSubmitSuiteLine();
            }
        }

        // Whether the suite went through on the model edited, as it is now.
        internal bool SuitePassedOnEditedModel =>
            this.EditedModelRow() is MainDocument.OscilloscopeRow model &&
            this.thisSuitePassedFingerprint is string passed &&
            passed == OscilloscopeModelValues.Fingerprint(model);

        internal void NoteSuitePassedForTests(OscilloscopeEntry model) => this.NoteSuitePassed(model);

        private void OnSubmitModelClick(object? sender, RoutedEventArgs e) => this.ShowSubmitPanel();

        private void OnCancelSendModelClick(object? sender, RoutedEventArgs e) => this.SubmitModelPanel.IsVisible = false;

        internal void ShowSubmitPanel()
        {
            if (this.EditedModelRow() is null)
            {
                return;
            }

            this.SubmitModelIntroText.Text = TabOscilloscope.SubmitIntro;
            this.ShowSubmitSuiteLine();
            this.FillSubmitAddress();

            if (string.IsNullOrWhiteSpace(this.SubmitModelNameTextBox.Text))
            {
                this.SubmitModelNameTextBox.Text = UserSettings.CreditName;
            }

            this.SubmitModelProblemText.IsVisible = false;
            this.SubmitModelPanel.IsVisible = true;
        }

        // Whether the suite went through on the model as it is now - said again after every change.
        internal void ShowSubmitSuiteLine()
        {
            bool passed = this.SuitePassedOnEditedModel;

            this.SubmitModelSuiteText.Text = passed ? TabOscilloscope.SuitePassedLine : TabOscilloscope.SuiteNotRunLine;
            this.SubmitModelSuiteText.Classes.Set("FieldProblem", !passed);
        }

        private void FillSubmitAddress()
        {
            ContactAddress address = ContactAddress.Choose(this.thisMaintainerAccount, UserSettings.ContactEmail, DateTimeOffset.UtcNow);

            this.SubmitModelEmailTextBox.Text = address.Email;
            this.SubmitModelEmailTextBox.IsReadOnly = address.IsFromAccount;
            this.SubmitModelEmailNoteText.Text = address.IsFromAccount
                ? "Your maintainer account's address, as you are signed in on the Maintainer tab."
                : "Used only to tell you whether your model was published.";
        }

        private async void OnSendModelClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole: an "async void" handler has no caller to take an exception.
            try
            {
                await this.SendModelAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnSendModelClick failed - [{ex}]");
                this.ShowSubmitProblems([$"The model could not be sent: {ex.Message}"]);
            }
        }

        internal async Task SendModelAsync()
        {
            if (this.EditedModelRowFollowingReload() is not MainDocument.OscilloscopeRow model)
            {
                return;
            }

            ContactAddress address = ContactAddress.Choose(this.thisMaintainerAccount, UserSettings.ContactEmail, DateTimeOffset.UtcNow);
            string email = address.IsFromAccount ? address.Email : this.SubmitModelEmailTextBox.Text?.Trim() ?? string.Empty;
            string name = this.SubmitModelNameTextBox.Text?.Trim() ?? string.Empty;
            string message = this.SubmitModelMessageTextBox.Text?.Trim() ?? string.Empty;
            IReadOnlyDictionary<string, string> values = OscilloscopeModelValues.Of(model);

            var problems = OscilloscopeModelRules.SubmissionProblems(model.Brand, model.SeriesOrModel, values, message).ToList();

            if (!address.IsFromAccount)
            {
                if (email.Length == 0)
                {
                    problems.Add(TabOscilloscope.NoAddress);
                }
                else if (!EmailAddressRules.IsPlausible(email))
                {
                    problems.Add(TabOscilloscope.BadAddress);
                }
            }

            if (problems.Count > 0)
            {
                this.ShowSubmitProblems(problems);
                return;
            }

            // The account's address goes with its token, as the Submit dialog sends it (code review,
            // 2026-10-29): a session the server has ended meanwhile makes no account there, and the
            // model is then taken as sent from the address - without one it was refused, and the
            // read-only box let nobody type one.
            var request = new OscilloscopeModelSubmissionRequest(
                model.Brand.Trim(),
                model.SeriesOrModel.Trim(),
                values,
                email,
                name.Length == 0 ? null : name,
                message.Length == 0 ? null : message,
                this.SuitePassedOnEditedModel);

            string? token = address.Account?.BearerToken;

            Func<OscilloscopeModelSubmissionRequest, string?, Task<OscilloscopeModelSendOutcome>> send =
                this.SendOverrideForTests ?? ((body, bearer) => new SubmissionClient().SubmitOscilloscopeModelAsync(body, bearerToken: bearer));

            WaitResult<OscilloscopeModelSendOutcome> waited = await BusyOverlay.RunAsync(
                this,
                TabOscilloscope.SendingModel,
                _ => send(request, token));

            if (waited.IsTimedOut)
            {
                this.ShowSubmitProblems([TabOscilloscope.SendNoAnswer]);
                return;
            }

            OscilloscopeModelSendOutcome outcome = waited.Value!;

            if (outcome.Answer is null)
            {
                this.ShowSubmitProblems(outcome.Problems);
                return;
            }

            if (!address.IsFromAccount)
            {
                UserSettings.ContactEmail = email;
            }

            UserSettings.SetOscilloscopeModelSent(model.Brand, model.SeriesOrModel, DateTimeOffset.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));

            this.SubmitModelPanel.IsVisible = false;
            this.SubmitModelMessageTextBox.Text = string.Empty;
            this.ShowModelSent(model, outcome.Answer.Message);
        }

        private void ShowSubmitProblems(IReadOnlyList<string> problems)
        {
            this.SubmitModelProblemText.Text = string.Join(Environment.NewLine, problems);
            this.SubmitModelProblemText.IsVisible = problems.Count > 0;
        }

        // "Sent in on <date>" under the editor's buttons - the server's own words right after sending.
        private void ShowModelSent(MainDocument.OscilloscopeRow model, string? justSaid = null)
        {
            string? date = UserSettings.GetOscilloscopeModelSent(model.Brand, model.SeriesOrModel);

            this.ModelSentText.Text = justSaid ?? (date is null ? string.Empty : $"Sent in for everybody on {date}. The outcome comes by mail.");
            this.ModelSentText.IsVisible = this.ModelSentText.Text.Length > 0;
        }
    }
}
