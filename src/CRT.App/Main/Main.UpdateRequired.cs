using System;
using System.Threading.Tasks;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;
using Handlers.OnlineHandling;

namespace CRT
{
    // ###########################################################################################
    // "CRT HAS TO BE UPDATED" ON THE DRAFTS AND MAINTAINER TABS (owner request, 2026-10-09: "When an
    // API diff requires for the CRT app to be updated, then for both the tabs, "Draft" and
    // "Maintainer" put a fullpage modal (or alike) in there, that cannot be closed, stating the app
    // needs to be updated").
    //
    // WHAT THIS PART OWNS: learning that the server turns this CRT away, and covering the tab(s) it
    // concerns with UpdateRequiredOverlay - which tab, in which words, is AppUpdateRequirement's.
    //
    //   - ONCE EITHER TAB IS IN USE, GET /api/health is asked for the server's API revision (fire and
    //     forget, like the check-in) - at launch, or the moment the Drafts tab appears or the
    //     Maintainer tab is turned on. A HIGHER one than this CRT's covers both tabs before anybody
    //     opens them. *** NOT FOR SOMEBODY WHO USES NEITHER (code review, 2026-10-09). *** Asked at
    //     every launch, every installation sent the server a request it counts (API usage) for two
    //     tabs it never shows - AppUpdateRequirement.AsksServer.
    //   - ANY 426 either client meets afterwards (ApiOutdatedSignal) covers the tab of the area that
    //     refused - a server deployed while CRT runs, or a minimum version - and the revision is
    //     asked again, so a revision that moved covers the other tab at once instead of the first
    //     time that tab asks.
    //
    // Only StartAsync listens to ApiOutdatedSignal and lets the question go out (tests build Main
    // but never start it - see the signal's header); OnWindowClosed lets go.
    //
    // *** THE OVERLAY'S BUTTON IS THE WAY OUT. *** It installs the update CRT's own check found -
    // the update banner's Install, which asks about unsaved table edits first - or, with none
    // found, opens the releases page (AppUpdateRequiredWording). A later update check that finds
    // one changes the button (CheckForAppUpdateNowAsync calls ApplyAppUpdateRequired).
    //
    // The update banner's own reasons (Main.Updates.cs) are untouched: the submission checks still
    // put the server's words in it, since it is seen from every tab.
    // ###########################################################################################
    public partial class Main
    {
        private readonly AppUpdateRequirement thisAppUpdateRequirement = new();

        // A health question in flight - a burst of refusals asks it once.
        private bool thisAskingApiRevision;

        // StartAsync has run, so the question may go out; and it has gone out once the tabs were
        // in use - from then on only a refusal asks it again.
        private bool thisMayAskApiRevision;
        private bool thisAskedApiRevisionInUse;

        // Stand-ins for the network and for what the button starts - see the test seams below.
        internal Func<Task<int?>>? ServerApiRevisionOverrideForTests { get; set; }

        internal Func<string?>? PendingVersionOverrideForTests { get; set; }

        internal Action<bool>? UpdateRequiredActionOverrideForTests { get; set; }

        // This CRT as the server reads it from the User-Agent - named in CRT's own sentence.
        private static CrtVersion? OwnVersion => CrtVersion.FromUserAgent(OnlineServices.VersionForServer);

        // ###########################################################################################
        // Called once from StartAsync: listen for refusals, and ask the revision if the tabs it is
        // about are in use already.
        // ###########################################################################################
        private void StartAppUpdateRequiredChecks()
        {
            ApiOutdatedSignal.Raised += this.OnApiOutdatedSignal;
            this.AllowApiRevisionQuestion();
        }

        private void AllowApiRevisionQuestion()
        {
            this.thisMayAskApiRevision = true;
            this.AskApiRevisionOnceInUse();
        }

        // StartAsync's half that asks, without listening to the static signal - for tests.
        internal void AllowApiRevisionQuestionForTests() => this.AllowApiRevisionQuestion();

        // ###########################################################################################
        // The revision asked the first time the Drafts tab is shown or the Maintainer tab turned on -
        // called from StartAsync and from both tabs' Apply...Visibility. Nothing before StartAsync,
        // and nothing more once asked: a refusal asks again by itself (ReportApiOutdated).
        // ###########################################################################################
        private void AskApiRevisionOnceInUse()
        {
            if (!this.thisMayAskApiRevision || this.thisAskedApiRevisionInUse)
                return;

            if (!AppUpdateRequirement.AsksServer(this.DraftsTabItem?.IsVisible == true, UserSettings.EnableMaintainerTab))
                return;

            this.thisAskedApiRevisionInUse = true;
            _ = this.AskServerApiRevisionAsync();
        }

        private void StopAppUpdateRequiredChecks() =>
            ApiOutdatedSignal.Raised -= this.OnApiOutdatedSignal;

        // Raised on whichever thread the answer arrived on.
        private void OnApiOutdatedSignal(AppUpdateArea area, string serversWords) =>
            Dispatcher.UIThread.Post(() => this.ReportApiOutdated(area, serversWords));

        // ###########################################################################################
        // A 426 from one area: its tab is covered in the server's words, and - unless both already
        // are - the revision is asked again.
        // ###########################################################################################
        internal void ReportApiOutdated(AppUpdateArea area, string serversWords)
        {
            if (this.thisAppUpdateRequirement.ApplyRefusal(area, serversWords, Main.OwnVersion))
                this.ApplyAppUpdateRequired();

            if (!this.thisAppUpdateRequirement.IsRequiredEverywhere)
                _ = this.AskServerApiRevisionAsync();
        }

        // ###########################################################################################
        // GET /api/health's revision against this CRT's. Every failure is no answer - the tabs stay
        // as they are, and a refusal will still cover its own tab when it comes.
        // ###########################################################################################
        internal async Task AskServerApiRevisionAsync()
        {
            if (this.thisAskingApiRevision)
                return;

            this.thisAskingApiRevision = true;

            try
            {
                int? serverRevision = this.ServerApiRevisionOverrideForTests is { } ask
                    ? await ask()
                    : await Main.AskHealthForApiRevisionAsync();

                if (this.thisAppUpdateRequirement.ApplyApiRevision(serverRevision, ClientVersionContract.ApiRevision, Main.OwnVersion))
                {
                    Logger.Warning($"The server serves API revision [{serverRevision}] and this CRT was built for [{ClientVersionContract.ApiRevision}] - CRT has to be updated before drafts can be submitted or the Maintainer tab used");
                    this.ApplyAppUpdateRequired();
                }
            }
            catch (Exception ex)
            {
                // No answer is not a refusal (offline, say) - nothing for anybody to act on.
                Logger.Info($"Asking the server for its API revision failed - [{ex.Message}]");
            }
            finally
            {
                this.thisAskingApiRevision = false;
            }
        }

        // The public health route, asked with no session (ReviewApiClient.GetServerVersionAsync).
        private static async Task<int?> AskHealthForApiRevisionAsync()
        {
            using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);

            ReviewApiResult<HealthStatus> answer = await client.GetServerVersionAsync();

            return answer.IsOk ? answer.Value!.ApiRevision : null;
        }

        // ###########################################################################################
        // Puts what AppUpdateRequirement says on each tab it covers - and says it again when the
        // words or the update on offer changed. A tab it does not cover is left alone.
        // ###########################################################################################
        internal void ApplyAppUpdateRequired()
        {
            string? pendingVersion = this.PendingVersionOverrideForTests is { } pending
                ? pending()
                : UpdateService.PendingVersion;

            if (this.thisAppUpdateRequirement.ReasonFor(AppUpdateArea.Drafts) is string draftsReason)
                this.TabDrafts.ShowUpdateRequired(AppUpdateRequiredWording.For(AppUpdateArea.Drafts, draftsReason, pendingVersion));

            if (this.thisAppUpdateRequirement.ReasonFor(AppUpdateArea.Maintainer) is string maintainerReason)
                this.TabMaintainer.ShowUpdateRequired(AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, maintainerReason, pendingVersion));
        }

        // ###########################################################################################
        // The overlay's button, on either tab: install the update found, or open the releases page.
        // ###########################################################################################
        private void OnUpdateRequiredAction(object? sender, EventArgs e)
        {
            if (sender is not UpdateRequiredOverlay { View: { } view })
                return;

            if (this.UpdateRequiredActionOverrideForTests is { } action)
            {
                action(view.ButtonInstalls);
                return;
            }

            if (view.ButtonInstalls)
                _ = this.InstallPendingUpdateAsync();
            else
                Main.OpenUrl(AppConfig.GitHubReleasesUrl);
        }
    }
}
