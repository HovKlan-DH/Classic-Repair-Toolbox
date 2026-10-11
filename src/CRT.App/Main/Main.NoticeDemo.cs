using Handlers.DataHandling;
using Handlers.OnlineHandling;
using Handlers.Theming;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "--demo-notices" (owner request, 2026-10-11: "can you demo all these different ways of
    // notifications, so I can see them all and how it behaves?"). After launch, every notice in
    // NoticeStack is played once in the real window, each step explained by the mode hint, so
    // their look and behaviour can be seen without waiting for each real cause. See Main.axaml.cs
    // for the file map, Main.Notices.cs for the notices themselves.
    //
    // The cards are the real ones, shown by their own code where that code needs no real cause,
    // so what is seen is what users get. The words are examples where the real ones need a real
    // cause (an update, a submission); the log says so (SimulationOptions.DescribeForLog). Nothing
    // the demo shows is remembered: the tooltip's "last check" is put back after each pretend check.
    //
    // A card ALREADY showing for a real reason is left alone - the demo must not reword it, and
    // closing a real card does real things (it remembers the BETA notice as read, for one).
    //
    // MainWindowTests runs the whole demo and fails if a card in the stack is never shown.
    // ###########################################################################################
    public partial class Main
    {
        // Waits between the demo's steps. Tests run it without waiting.
        internal Func<TimeSpan, Task>? NoticeDemoDelayOverrideForTests { get; set; }

        private const int NoticeDemoSteps = 8;

        internal async Task PlayNoticeDemoAsync()
        {
            try
            {
                Func<TimeSpan, Task> wait = this.NoticeDemoDelayOverrideForTests ?? (delay => Task.Delay(delay));
                string source = AppConfig.GetOnlineSourceLabel();

                await wait(TimeSpan.FromSeconds(2));

                this.NarrateNoticeDemo(1, this.CanPretendDataCheck()
                    ? "A data check, as at every launch. While it runs, its card says what it is doing. " +
                      "Nothing changed, so it says so - and goes by itself after 8 seconds."
                    : Main.NoticeDemoDataCheckLeftAlone);
                await this.PretendDataCheckAsync(wait, source, changedCount: 0);
                await wait(Main.NoticeGoesAfter + TimeSpan.FromSeconds(2));

                this.NarrateNoticeDemo(2, this.CanPretendDataCheck()
                    ? "A data check that changed files. This card stays until you refresh the board or close it."
                    : Main.NoticeDemoDataCheckLeftAlone);
                await this.PretendDataCheckAsync(wait, source, changedCount: 3);
                await wait(TimeSpan.FromSeconds(5));

                this.NarrateNoticeDemo(3,
                    "A newer CRT to install. Each new card goes on top of the others - nothing else on screen moves.");
                if (!this.UpdateBanner.IsVisible)
                {
                    this.ShowApplicationUpdateAvailableBanner();

                    // No update found (nor simulated): an example version. Install then fails as it
                    // would for real - "could not be installed" - which is safe to try.
                    if (UpdateService.PendingVersion == null)
                    {
                        NoticeEmphasis.Show(this.UpdateBannerText, "Version [99.0.0] is available");
                    }
                }
                await wait(TimeSpan.FromSeconds(5));

                this.NarrateNoticeDemo(4,
                    "Board data that needs a newer CRT than this one.");
                if (!this.MainExcelRequiresAppUpdateBanner.IsVisible)
                {
                    this.ShowMainExcelRequiresAppUpdateBanner();
                }
                await wait(TimeSpan.FromSeconds(5));

                this.NarrateNoticeDemo(5,
                    "The published data for the board on screen changed after you started a draft of it. " +
                    "\"What changed\" opens the report when there is a draft.");
                if (!this.SyncBanner.IsVisible)
                {
                    this.ShowDraftDriftBanner();
                }
                await wait(TimeSpan.FromSeconds(5));

                this.NarrateNoticeDemo(6,
                    "Your submission reached the stable source while you download from the BETA source.");
                if (!this.SourceSwitchBanner.IsVisible)
                {
                    NoticeEmphasis.Show(
                        this.SourceSwitchBannerText,
                        "Commodore / C64 / 250407 is now published to the stable source. You are downloading data from " +
                        "the BETA source - you can switch back to the stable source on the Configuration tab.");
                    this.SourceSwitchBanner.IsVisible = true;
                }
                await wait(TimeSpan.FromSeconds(5));

                this.NarrateNoticeDemo(7,
                    "The label editor holds the board. Its card has no close button - it goes when the editor " +
                    "closes; here, after 8 seconds.");
                if (!this.SchematicsLabelEditorModeBanner.IsVisible)
                {
                    this.SetSchematicsLabelEditorModeBannerVisible(true);
                    await wait(TimeSpan.FromSeconds(8));
                    this.SetSchematicsLabelEditorModeBannerVisible(false);
                }

                this.NarrateNoticeDemo(8,
                    "That is all of them. A card that has just appeared lets clicks through for half a second, " +
                    "so a click meant for what was under it is not lost. In a small window the cards keep to the " +
                    "right half. Close the rest with their X.");
                await wait(TimeSpan.FromSeconds(20));
                this.HideModeHint();
            }
            catch (Exception ex)
            {
                Logger.Warning($"The notice demo stopped - [{ex.Message}]");
            }
        }

        private void NarrateNoticeDemo(int step, string text) =>
            this.ShowModeHint($"Notice demo {step} of {Main.NoticeDemoSteps}: {text}");

        private const string NoticeDemoDataCheckLeftAlone =
            "The data-check card is showing a real check, so the demo leaves it alone.";

        // What the demo's last pretend check left on the data-check card - so the next one can tell
        // its own card from a real one.
        private string? thisNoticeDemoDataCheckText;

        // ###########################################################################################
        // Whether the data-check card is free for a pretend check: not showing a real check running,
        // nor a real check's outcome (code review, 2026-10-11 - the demo overwrote a real "please
        // refresh board" with its progress, then "all up to date", which went by itself). Its own
        // card from the step before is free.
        // ###########################################################################################
        private bool CanPretendDataCheck() =>
            !this.thisIsDataCheckOnCard &&
            (!this.DataCheckBanner.IsVisible ||
             string.Equals(NoticeEmphasis.TextOf(this.DataCheckBannerText), this.thisNoticeDemoDataCheckText, StringComparison.Ordinal));

        // ###########################################################################################
        // A check through the real card and spinner, with made-up files. The tooltip's "last check"
        // is put back afterwards - a pretend check must not claim a real one happened. Nothing at all
        // while the card shows a real check (CanPretendDataCheck).
        // ###########################################################################################
        private async Task PretendDataCheckAsync(Func<TimeSpan, Task> wait, string source, int changedCount)
        {
            if (!this.CanPretendDataCheck())
            {
                return;
            }

            string? lastRealCheck = this.thisLastDataSyncCheck;
            string[] files =
            [
                "Commodore/C64/250407/C64 250407 - schematic 1 of 2.png",
                "Commodore/C64/250407/C64 250407 - schematic 2 of 2.png",
                "Commodore/C64/250407/C64 250407 - PCB top.png",
            ];

            this.BeginDataSyncProgress();
            this.StartDataSyncStatusIconSpin();

            try
            {
                this.ReportDataSyncStatus("Fetching online file manifest...");
                await wait(TimeSpan.FromSeconds(1));

                for (int i = 0; i < files.Length; i++)
                {
                    this.ReportDataSyncStatus($"Downloading file [{i + 1}] of [{files.Length}] from the {source}:");
                    this.ReportDataSyncFile(files[i]);
                    await wait(TimeSpan.FromSeconds(1));
                }
            }
            finally
            {
                this.StopDataSyncStatusIconSpin();
                this.FinishDataCheck(changedCount, atLaunch: true);
                this.thisLastDataSyncCheck = lastRealCheck;
                this.thisNoticeDemoDataCheckText = NoticeEmphasis.TextOf(this.DataCheckBannerText);
                this.RefreshDataSyncStatusTip();
            }
        }
    }
}
