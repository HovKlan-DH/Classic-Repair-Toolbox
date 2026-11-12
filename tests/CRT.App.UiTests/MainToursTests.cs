using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;
using Handlers.Tours;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The introduction tours in CRT's window (Main.Tours.cs; owner request, 2026-11-11). What must hold: every control a tour
// frames is in the window; no tour starts before StartAsync, which no test runs, opens the gate; a tab's tour shows
// once, and is remembered whether taken to its end or skipped; one seen at an older version shows what came since;
// "--show-tours" shows them whatever was seen and remembers nothing; "Take the introduction tours again" forgets them.
//
// As MainBoardSwitchTests: the real Main with a real board file, StartAsync never called - the gate is opened by
// LetToursStartForTests, and OfferTourNow asked directly, never through the timers that wait for a tab to settle.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainToursTests : IDisposable
{
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private readonly TempWorkspace thisWorkspace = new();

    public MainToursTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    // A data root listing one board, with a picture and a component - in PAL and NTSC when asked, so the region buttons
    // show, or a board file that cannot be read.
    private void ListBoard(bool withRegions = false, bool unreadable = false)
    {
        string dataRoot = this.thisWorkspace.Path_("Data");
        var listed = new MainFiles.Board("Commodore/C64/250407", "Commodore 64", "250407");
        DataManager.LoadFrom(dataRoot, MainFiles.Write(dataRoot, listed));

        BoardDocument.StoredFileRow picture = StoredFiles.Put(dataRoot, "Images/Top.png", MainToursTests.Png);

        StoredFiles.WriteBoard(
            DraftBoardSource.PublishedPathOf(dataRoot, listed.Identity),
            new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = picture.Name }],
                Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
                Regions = withRegions ? ["PAL", "NTSC"] : [],
            },
            picture);

        if (unreadable)
        {
            File.WriteAllText(DraftBoardSource.PublishedPathOf(dataRoot, listed.Identity), "{ this is no board");
        }
    }

    // CRT's window shown on the board, its first board drawn, the gate open unless asked not to.
    private async Task WithMainAsync(Func<CRT.Main, Task> body, bool openGate = true, bool withRegions = false, bool unreadableBoard = false)
    {
        this.ListBoard(withRegions, unreadableBoard);

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main { Width = 1400, Height = 900 };
            main.Show();
            main.PopulateHardwareDropDown();

            try
            {
                for (int attempt = 0; attempt < 1000 && !main.FirstBoardShown.IsCompleted; attempt++)
                {
                    Dispatcher.UIThread.RunJobs();
                    await Task.Delay(10);
                }

                Assert.True(main.FirstBoardShown.IsCompleted, "the board never came");
                MainToursTests.Settle(main);

                if (openGate)
                {
                    main.LetToursStartForTests();
                }

                await body(main);
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    private static void Settle(CRT.Main main)
    {
        for (int pass = 0; pass < 3; pass++)
        {
            Dispatcher.UIThread.RunJobs();
            main.UpdateLayout();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
    }

    private static void OnTab(CRT.Main main, string tabItem)
    {
        main.GetControl<TabControl>("MainTabControl").SelectedItem = main.GetControl<TabItem>(tabItem);
        MainToursTests.Settle(main);
    }

    // ###########################################################################################
    // The tab's tour offered as the window's timer offers it: once nothing holds it back. A moment of the tab settling - a
    // wait passing as the "Contribute" tab is entered, a popup closing - is waited out, as the timer waits it out a second
    // at a time; asked at once, the tour was refused for it once in a full run (2026-11-12). What held it back is the
    // assertion's message (TourWaitReasonForTests).
    // ###########################################################################################
    private static bool Offered(CRT.Main main)
    {
        for (int attempt = 0; attempt < 200 && main.TourMustWait(); attempt++)
        {
            Thread.Sleep(10);
            Dispatcher.UIThread.RunJobs();
        }

        return main.OfferTourNow();
    }

    private static TourOverlay Overlay(CRT.Main main) => main.GetControl<TourOverlay>("TourOverlay");

    private static string? Kicker(CRT.Main main) => Overlay(main).GetControl<TextBlock>("KickerText").Text;

    private static int StepCount(CRT.Main main) => Overlay(main).GetControl<StackPanel>("ProgressDots").Children.Count;

    // To the step shown, as a press of "Next" does - so the dots show the count.
    private static void ToFirstStep(CRT.Main main)
    {
        Overlay(main).Next();
        MainToursTests.Settle(main);
    }

    // On, as "Next" goes, to the step of `step`'s headline - by its words, as a step not on screen is left out of the
    // showing and the catalogue's index is then not the one shown (the trace step on a board without KiCad data).
    private static void ToStep(CRT.Main main, TourStep step)
    {
        TourOverlay overlay = Overlay(main);

        while (overlay.IsShowing && overlay.GetControl<TextBlock>("HeadlineText").Text != step.Headline)
        {
            overlay.Next();
        }

        Assert.True(overlay.IsShowing, $"the step [{step.Headline}] was not shown");
        MainToursTests.Settle(main);
    }

    // ###########################################################################################
    // *** EVERY CONTROL A TOUR FRAMES IS IN CRT'S WINDOW. *** A step names its controls by x:Name, which the compiler
    // never sees: a control renamed or removed would leave its step out of every showing, silently.
    // ###########################################################################################
    [Fact]
    public void Every_control_a_tour_frames_is_in_CRTs_window()
    {
        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            try
            {
                var missing = TourCatalogue.All
                    .SelectMany(tour => tour.Steps.SelectMany(step => step.Targets.Select(path => (tour.Tab, path))))
                    .Where(target => main.FindTourTarget(target.path) is null)
                    .Select(target => $"{target.Tab}: {target.path}")
                    .ToList();

                Assert.Empty(missing);
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // Only StartAsync opens the gate - and no test runs it, so no other test of CRT's window ever meets a tour.
    [Fact]
    public async Task No_tour_starts_before_the_gate_is_open()
    {
        await this.WithMainAsync(main =>
        {
            Assert.False(main.OfferTourNow());
            Assert.False(Overlay(main).IsVisible);
            return Task.CompletedTask;
        }, openGate: false);
    }

    // ###########################################################################################
    // *** EVERYBODY SEES A TAB'S INTRODUCTION ONCE (owner decision, 2026-11-11: "Everyone should see it"). *** Nothing
    // stored: the "Schematics" tab's introduction, every step on screen. Taken to its end, it is remembered as the
    // tour's version, and not offered again.
    // ###########################################################################################
    [Fact]
    public async Task The_schematics_tabs_introduction_shows_once_and_is_remembered()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());

            TourOverlay overlay = Overlay(main);
            Assert.True(overlay.IsVisible);
            Assert.Equal("Introduction tour", Kicker(main));

            ToFirstStep(main);

            // Every step but the trace step - this board has no KiCad data, so its badge is not there - and, the first
            // introduction taken, the card on the other tabs.
            Assert.Equal(TourCatalogue.Of(TourTab.Schematics).Steps.Count(step => !step.Targets.Contains("TabSchematicsControl/KiCadDataBadge")) + 1, StepCount(main));

            // The first step frames the component lists and the filter box under them.
            Rect hole = Assert.NotNull(overlay.HoleShown);
            Control filter = main.GetControl<TextBox>("ComponentSearchTextBox");
            Assert.True(hole.Contains(new Rect(filter.TranslatePoint(new Point(0, 0), overlay)!.Value, filter.Bounds.Size)));

            while (overlay.IsShowing)
            {
                overlay.Next();
            }

            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Schematics"));
            Assert.False(main.OfferTourNow());
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** THE PANELS' STEP SHOWS THEM EXPANDED, AND PUTS THEM BACK (owner request, 2026-11-11: "add an addition step ...
    // to show the 3 panels. Please show them expanded"). *** "Labels visible" folded away by its owner opens for the step,
    // framed whole, and folds again as the step is left - by "Next", and by "Skip tour" - its stored setting never written.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task The_panels_step_shows_them_expanded_and_puts_them_back(bool leftBySkipping)
    {
        UserSettings.SchematicsLabelsPanelExpanded = false;

        await this.WithMainAsync(main =>
        {
            StackPanel labels = main.TabSchematicsControl.GetControl<StackPanel>("LabelsListPanel");
            Assert.False(labels.IsVisible);

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            TourOverlay overlay = Overlay(main);
            ToStep(main, TourCatalogue.Of(TourTab.Schematics).Steps.Single(step => step.Stage == TourStage.SchematicsPanelsExpanded));
            Assert.True(labels.IsVisible);
            Assert.Equal("Collapse", main.TabSchematicsControl.GetControl<Button>("ToggleLabelsPanelButton").Content);

            Control panel = main.TabSchematicsControl.GetControl<Border>("LabelsPanel");
            Rect hole = Assert.NotNull(overlay.HoleShown);
            Assert.True(hole.Contains(new Rect(panel.TranslatePoint(new Point(0, 0), overlay)!.Value, panel.Bounds.Size)), $"the hole {hole} does not hold the expanded panel");

            if (leftBySkipping)
            {
                overlay.Skip();
            }
            else
            {
                overlay.Next();
            }

            MainToursTests.Settle(main);
            Assert.False(labels.IsVisible);
            Assert.Equal("Expand", main.TabSchematicsControl.GetControl<Button>("ToggleLabelsPanelButton").Content);
            Assert.False(UserSettings.SchematicsLabelsPanelExpanded);
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** A STEP SAYS NOTHING OF WHAT IS NOT ON SCREEN (owner question, 2026-11-11: "Will this step always show the
    // current board selected, as that may not include the regions buttons? Then what? Same for step 4"). *** This board
    // has no regions and no KiCad data: the buttons' step names no "PAL", the panels' step no "Global settings" and no
    // "Important KiCad signals" - and both say the rest.
    // ###########################################################################################
    [Fact]
    public async Task A_step_says_nothing_of_a_control_the_board_does_not_show()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            TourOverlay overlay = Overlay(main);
            TextBlock body = overlay.GetControl<TextBlock>("BodyText");
            IReadOnlyList<TourStep> steps = TourCatalogue.Of(TourTab.Schematics).Steps;

            ToStep(main, steps.Single(step => step.Targets.Contains("RegionButtonsGrid")));
            Assert.Contains("Mark all", BoldRuns.TextOf(body), StringComparison.Ordinal);
            Assert.DoesNotContain("PAL", BoldRuns.TextOf(body), StringComparison.Ordinal);

            ToStep(main, steps.Single(step => step.Stage == TourStage.SchematicsPanelsExpanded));

            Assert.Contains("Labels visible", BoldRuns.TextOf(body), StringComparison.Ordinal);
            Assert.DoesNotContain("Global settings", BoldRuns.TextOf(body), StringComparison.Ordinal);
            Assert.DoesNotContain("Important KiCad signals", BoldRuns.TextOf(body), StringComparison.Ordinal);

            overlay.Skip();
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** WHAT A TOUR LEFT OUT IS REMEMBERED, AND COMES BACK WHOLE (code review, 2026-11-12: the trace step and "Export to
    // PDF" never reached a new user whose first look had no KiCad data or no workbook). *** On a board with no regions and
    // no KiCad data, the introduction ends remembering the steps it said in part or not at all - and offers nothing more
    // there, as none of them can be said whole.
    // ###########################################################################################
    [Fact]
    public async Task What_the_introduction_left_out_is_remembered()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            TourOverlay overlay = Overlay(main);

            while (overlay.IsShowing)
            {
                overlay.Next();
            }

            IReadOnlyList<string> leftOut = UserSettings.IntroductionTourStepsLeftOut("Schematics");
            Assert.Contains("Below the component list", leftOut);
            Assert.Contains("Highlight a trace", leftOut);
            Assert.Contains("Panels on the schematic image", leftOut);
            Assert.DoesNotContain("Find a component", leftOut);

            Assert.False(main.OfferTourNow());
            return Task.CompletedTask;
        });
    }

    // On a board WITH regions, the step about the buttons under the list comes back whole - "PAL" said, no "New in"
    // badge - while the trace step, still not on screen, waits on. Taken, it is left out no more.
    [Fact]
    public async Task A_step_left_out_comes_back_whole_once_it_is_on_screen()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0", ["Below the component list", "Highlight a trace"]);
        UserSettings.SetIntroductionTourSeen("Contribute", "4.0.0");

        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            TourOverlay overlay = Overlay(main);

            Assert.Equal("Introduction tour, continued", Kicker(main));
            Assert.Contains("left out", BoldRuns.TextOf(overlay.GetControl<TextBlock>("BodyText")), StringComparison.Ordinal);

            ToFirstStep(main);
            Assert.Equal(1, StepCount(main));
            Assert.Equal("Below the component list", overlay.GetControl<TextBlock>("HeadlineText").Text);
            Assert.Contains("PAL", BoldRuns.TextOf(overlay.GetControl<TextBlock>("BodyText")), StringComparison.Ordinal);
            Assert.False(overlay.GetControl<Border>("NewBadge").IsVisible);

            overlay.Next();
            Assert.False(overlay.IsShowing);

            Assert.Equal(["Highlight a trace"], UserSettings.IntroductionTourStepsLeftOut("Schematics"));
            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Schematics"));
            Assert.False(main.OfferTourNow());
            return Task.CompletedTask;
        }, withRegions: true);
    }

    // ###########################################################################################
    // A closed window offers no tour: the timer a tab's tour waits on goes with it - left, it looked again every second
    // for up to two minutes on the tests' shared dispatcher, keeping the closed window alive (code review, 2026-11-12).
    // ###########################################################################################
    [Fact]
    public async Task Closing_the_window_stops_the_tour_waiting_to_be_offered()
    {
        CRT.Main? closed = null;

        await this.WithMainAsync(main =>
        {
            OnTab(main, "FeedbackTabItem");
            Assert.True(main.IsTourTimerRunningForTests);

            closed = main;
            return Task.CompletedTask;
        });

        Assert.False(Assert.IsType<CRT.Main>(closed).IsTourTimerRunningForTests);
    }

    // A tool of the schematic image on - "Draw trace" here - holds the tour back: the Escape that skips it would end the
    // tool too (code review, 2026-11-12).
    [Fact]
    public async Task A_tour_waits_while_a_tool_of_the_schematic_image_is_on()
    {
        await this.WithMainAsync(main =>
        {
            main.TabSchematicsControl.SetDrawingTraces(true);
            Assert.True(main.TabSchematicsControl.IsDrawingTraces);
            Assert.True(main.TourMustWait());
            Assert.False(main.OfferTourNow());

            main.TabSchematicsControl.SetDrawingTraces(false);
            Assert.False(main.TourMustWait());
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // A tour SKIPPED was turned down: what it left out does not come back (code review, 2026-11-12: skipped on its first
    // card, it came back later as "what an earlier tour left out") - and what an earlier one left out goes with it.
    // ###########################################################################################
    [Fact]
    public async Task A_tour_skipped_leaves_nothing_to_come_back()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Overlay(main).Skip();

            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Schematics"));
            Assert.Empty(UserSettings.IntroductionTourStepsLeftOut("Schematics"));
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task A_tour_of_what_was_left_out_skipped_forgets_it()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0", ["Below the component list", "Highlight a trace"]);

        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Overlay(main).Skip();

            Assert.Empty(UserSettings.IntroductionTourStepsLeftOut("Schematics"));
            Assert.False(main.OfferTourNow());
            return Task.CompletedTask;
        }, withRegions: true);
    }

    // ###########################################################################################
    // Steps left out that cannot be said whole now - the trace step on a board with no KiCad data - are nothing to wait
    // for: the timer asks once and stops, while a tool is on too (code review, 2026-11-12: it looked again every second
    // for two minutes at every visit).
    // ###########################################################################################
    [Fact]
    public async Task Steps_left_out_that_cannot_be_said_whole_are_not_waited_for()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0", ["Highlight a trace"]);
        UserSettings.SetIntroductionTourSeen("Contribute", "4.0.0");

        await this.WithMainAsync(async main =>
        {
            OnTab(main, "ContributeTabItem");
            OnTab(main, "SchematicsTabItem");
            main.TabSchematicsControl.SetDrawingTraces(true);
            Assert.True(main.TourMustWait());
            Assert.True(main.IsTourTimerRunningForTests);

            for (int attempt = 0; attempt < 100 && main.IsTourTimerRunningForTests; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(20);
            }

            Assert.False(main.IsTourTimerRunningForTests, "the timer kept waiting for steps that cannot be said whole");
            Assert.False(Overlay(main).IsShowing);
            Assert.Equal(["Highlight a trace"], UserSettings.IntroductionTourStepsLeftOut("Schematics"));
            main.TabSchematicsControl.SetDrawingTraces(false);
        });
    }

    // A picture tool asked for and waiting - here for the board's KiCad data - holds the tour back as one that is on: it
    // would start under the tour (code review, 2026-11-12).
    [Fact]
    public async Task A_tour_waits_for_a_picture_tool_waiting_to_start()
    {
        await this.WithMainAsync(main =>
        {
            main.TabSchematicsControl.KiCadLoadingForTests();
            main.StartPictureTool(PictureTool.AlignKiCad, "Top", readBoardFirst: false);

            Assert.True(main.TabSchematicsControl.IsPictureToolPending);
            Assert.False(main.IsPictureToolOn);
            Assert.True(main.TourMustWait());
            Assert.False(main.OfferTourNow());

            main.TabSchematicsControl.KiCadSettled(hasProject: false);
            Assert.False(main.TabSchematicsControl.IsPictureToolPending);
            return Task.CompletedTask;
        });
    }

    // A drop-down list open in CRT's window holds the tour back: it stands over the tour, and a click there would pick a
    // board under a tour that takes no clicks (code review, 2026-11-12).
    [Fact]
    public async Task A_tour_waits_while_a_drop_down_list_is_open()
    {
        await this.WithMainAsync(main =>
        {
            ComboBox hardware = main.GetControl<ComboBox>("HardwareComboBox");
            hardware.IsDropDownOpen = true;
            MainToursTests.Settle(main);

            Assert.True(main.TourMustWait());
            Assert.False(main.OfferTourNow());

            hardware.IsDropDownOpen = false;
            MainToursTests.Settle(main);
            Assert.False(main.TourMustWait());
            return Task.CompletedTask;
        });
    }

    // And a flyout - a color picker's - placed at a control of the window, though it stands in no part of the window's
    // tree (code review, 2026-11-12: a color could be picked and saved under the tour).
    [Fact]
    public async Task A_tour_waits_while_a_flyout_is_open()
    {
        await this.WithMainAsync(main =>
        {
            var flyout = new Flyout { Content = new TextBlock { Text = "A flyout" } };
            flyout.ShowAt(main.GetControl<ComboBox>("HardwareComboBox"));
            MainToursTests.Settle(main);

            Assert.True(main.TourMustWait());

            flyout.Hide();
            MainToursTests.Settle(main);
            Assert.False(main.TourMustWait());
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // Coming back to the front offers only a tour whose wait ran out while the window was NOT in front (code review,
    // 2026-11-12: offered at every activation, a tour of what an earlier one left out came up in the middle of work as a
    // component's window closed). Here one that can be said whole is waiting - and still only the held-back one is asked.
    // ###########################################################################################
    [Fact]
    public async Task Coming_back_to_the_front_offers_only_a_tour_held_back_by_it()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0", ["Below the component list"]);

        await this.WithMainAsync(main =>
        {
            main.ComeBackToTheFrontForTests();
            Assert.False(main.IsTourTimerRunningForTests);

            main.TourWaitedOutBehindForTests();
            main.ComeBackToTheFrontForTests();
            Assert.True(main.IsTourTimerRunningForTests);
            return Task.CompletedTask;
        }, withRegions: true);
    }

    // ###########################################################################################
    // A wait that runs out shows nothing - shown anyway, the "Schematics" introduction came over KiCad data still being
    // read and left out its steps (code review, 2026-11-12). It comes the next time the tab is shown.
    // ###########################################################################################
    [Fact]
    public async Task A_wait_that_runs_out_shows_no_tour()
    {
        await this.WithMainAsync(async main =>
        {
            main.TourWaitsAtMost = 0;
            OnTab(main, "ContributeTabItem");
            OnTab(main, "SchematicsTabItem");
            main.TabSchematicsControl.KiCadLoadingForTests();

            for (int attempt = 0; attempt < 100 && main.IsTourTimerRunningForTests; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(20);
            }

            Assert.False(main.IsTourTimerRunningForTests);
            Assert.False(Overlay(main).IsShowing);
            Assert.Null(UserSettings.IntroductionTourSeen("Schematics"));

            main.TabSchematicsControl.KiCadSettled(hasProject: false);
            Assert.True(Offered(main), main.TourWaitReasonForTests());
        });
    }

    // ###########################################################################################
    // ON SCREEN IS IN VIEW: a control scrolled out of its page is not on screen, and a step about it frames nothing there
    // (code review, 2026-11-12: framed outside the page, over the tab strip). The "Configuration" tab's last folder button,
    // below the fold of a low window - until it is scrolled into view.
    // ###########################################################################################
    [Fact]
    public async Task A_control_scrolled_out_of_its_page_is_not_on_screen()
    {
        await this.WithMainAsync(main =>
        {
            main.Height = 500;
            OnTab(main, "ConfigurationTabItem");

            const string button = "TabConfiguration/OpenLogsFolderButton";
            Assert.NotNull(main.FindTourTarget(button));
            Assert.False(main.IsTourTargetOnScreen(button));

            main.FindTourTarget(button)!.BringIntoView();
            MainToursTests.Settle(main);
            Assert.True(main.IsTourTargetOnScreen(button));
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // A TOUR WAITS FOR THE PERSON TO PAUSE TYPING (code review, 2026-11-12: one that had waited for KiCad data opened in the
    // middle of typing in the filter, and an Escape meant for the filter skipped the introduction for good).
    // ###########################################################################################
    [Fact]
    public async Task A_tour_waits_while_the_person_types()
    {
        await this.WithMainAsync(main =>
        {
            main.KeyPress(Key.U, RawInputModifiers.None, PhysicalKey.U, keySymbol: "u");
            main.KeyRelease(Key.U, RawInputModifiers.None, PhysicalKey.U, keySymbol: "u");
            Dispatcher.UIThread.RunJobs();

            Assert.True(main.TourMustWait());
            Assert.False(main.OfferTourNow());

            main.LastKeyForTourWasAgoForTests(CRT.Main.TourAfterTyping + TimeSpan.FromMilliseconds(100));
            Assert.False(main.TourMustWait());
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            return Task.CompletedTask;
        });
    }

    // While a tour shows, the focus is its own - a click on it does not send the focus back to the component filter
    // (code review, 2026-11-12).
    [Fact]
    public async Task While_a_tour_shows_the_focus_is_not_sent_back_to_the_filter()
    {
        await this.WithMainAsync(main =>
        {
            // The focus on a drop-down list: a click would send it back to the filter.
            Assert.True(main.GetControl<ComboBox>("HardwareComboBox").Focus());
            Assert.True(main.ShouldReturnFocusToComponentSearch());

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            MainToursTests.Settle(main);

            Assert.False(main.ShouldReturnFocusToComponentSearch());
            Assert.True(Overlay(main).GetControl<Button>("NextButton").IsFocused);
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // A board load that throws ends its wait for KiCad data, as an early return does (code review, 2026-11-12: left "still
    // loading" until another board, the "Schematics" tour polled at every visit, "Align KiCad" sat at "Opening...", and
    // "Loading KiCad data..." stood on). On the UI thread right after the viewer's reset, and in the background half.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task A_board_load_that_throws_ends_its_wait_for_KiCad_data(bool inBackground)
    {
        await this.WithMainAsync(async main =>
        {
            TabSchematics tab = main.TabSchematicsControl;
            Action fail = () =>
            {
                tab.SetKiCadInitializingIndicatorVisible(true);
                throw new InvalidOperationException("a load step failed - a test's");
            };

            if (inBackground)
                main.ThrowInBackgroundLoadForTests = fail;
            else
                main.ThrowAfterViewerResetForTests = fail;

            main.ReloadCurrentBoardFromDisk("Top");

            for (int attempt = 0; attempt < 200 && (tab.IsKiCadStillLoading || tab.KiCadInitializingIndicator.IsVisible); attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.False(tab.IsKiCadStillLoading);
            Assert.False(tab.KiCadInitializingIndicator.IsVisible);

            main.ThrowInBackgroundLoadForTests = null;
            main.ThrowAfterViewerResetForTests = null;
        });
    }

    // A minimized window holds the tour back too - offered then, it was dropped until another tab was picked (code review,
    // 2026-11-12).
    [Fact]
    public async Task A_tour_waits_while_the_window_is_minimized()
    {
        await this.WithMainAsync(main =>
        {
            main.WindowState = WindowState.Minimized;
            Assert.True(main.TourMustWait());
            Assert.False(main.OfferTourNow());

            main.WindowState = WindowState.Normal;
            main.Activate();
            Dispatcher.UIThread.RunJobs();
            Assert.False(main.TourMustWait());
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // A tab whose tour was seen, with nothing new or left out, is not waited for (code review, 2026-11-12: every tab picked
    // looked again every second for up to two minutes while the board's KiCad data was read). The timer asks once, finds
    // nothing to offer, and stops.
    // ###########################################################################################
    [Fact]
    public async Task A_tab_with_nothing_to_offer_is_not_waited_for()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0");
        UserSettings.SetIntroductionTourSeen("Contribute", "4.0.0");

        await this.WithMainAsync(async main =>
        {
            OnTab(main, "ContributeTabItem");
            OnTab(main, "SchematicsTabItem");
            main.TabSchematicsControl.KiCadLoadingForTests();
            Assert.True(main.TourMustWait());
            Assert.True(main.IsTourTimerRunningForTests);

            for (int attempt = 0; attempt < 100 && main.IsTourTimerRunningForTests; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(20);
            }

            Assert.False(main.IsTourTimerRunningForTests, "the timer kept looking for a tour there is none of");
            Assert.False(Overlay(main).IsShowing);
        });
    }

    // ###########################################################################################
    // A board that cannot be read has no KiCad data to come: the "Schematics" tour does not wait for it (code review,
    // 2026-11-12: the load's early returns left it "still loading", and the tour came only after two minutes).
    // ###########################################################################################
    [Fact]
    public async Task The_schematics_tour_does_not_wait_for_KiCad_data_of_a_board_that_cannot_be_read()
    {
        await this.WithMainAsync(main =>
        {
            Assert.False(main.TabSchematicsControl.IsKiCadStillLoading);
            Assert.False(main.TourMustWait());
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            return Task.CompletedTask;
        }, unreadableBoard: true);
    }

    // The "Schematics" tour waits for the board's KiCad data - offered first, the panels' step found "Global settings" and
    // "Important KiCad signals" not there yet (2026-11-11, in a render of the C64 250407). Another tab does not wait for it.
    [Fact]
    public async Task The_schematics_tour_waits_for_the_boards_KiCad_data()
    {
        await this.WithMainAsync(main =>
        {
            Assert.False(main.TourMustWait());

            main.TabSchematicsControl.KiCadLoadingForTests();
            Assert.True(main.TourMustWait());

            OnTab(main, "ContributeTabItem");
            Assert.False(main.TourMustWait());

            OnTab(main, "SchematicsTabItem");
            main.TabSchematicsControl.KiCadSettled(hasProject: false);
            Assert.False(main.TourMustWait());
            return Task.CompletedTask;
        });
    }

    // Skipped - "Skip tour" or Escape - is seen too: offered again, it would come back at every launch.
    [Fact]
    public async Task A_tour_skipped_with_Escape_is_remembered_too()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());

            main.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);
            main.KeyRelease(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.False(Overlay(main).IsShowing);
            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Schematics"));
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** AFTER AN UPDATE, WHAT CAME SINCE (owner question, 2026-11-11: "How to handle updates, and then new
    // functionalities?"). *** Seen at a version before every step: "Tour of what's new", each step badged with the version
    // it came with. Seen at the newest: nothing.
    // ###########################################################################################
    [Fact]
    public async Task A_tour_seen_at_an_older_version_shows_what_came_since()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "3.1.1");

        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Assert.Equal("Tour of what's new", Kicker(main));

            ToFirstStep(main);
            Assert.Equal("New in 4.0.0", Overlay(main).GetControl<TextBlock>("NewBadgeText").Text);

            Overlay(main).Skip();
            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Schematics"));
            Assert.False(main.OfferTourNow());
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** "--show-tours" (owner request, 2026-11-11: "I need a possibility to enforce this every time when in
    // implementation mode"). *** Seen or not, the tab's tour shows - once in the run - and nothing is remembered.
    // ###########################################################################################
    [Fact]
    public async Task Forced_tours_show_whatever_was_seen_and_remember_nothing()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0");

        await this.WithMainAsync(main =>
        {
            main.TourOptions = SimulationOptions.Parse(["--show-tours"]);

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Assert.Equal("Introduction tour", Kicker(main));

            Overlay(main).Skip();

            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Schematics"));
            Assert.False(main.OfferTourNow());
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task Forced_with_a_version_shows_what_came_since_it()
    {
        await this.WithMainAsync(main =>
        {
            main.TourOptions = SimulationOptions.Parse(["--show-tours=3.1.1"]);

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Assert.Equal("Tour of what's new", Kicker(main));

            Overlay(main).Skip();
            Assert.Null(UserSettings.IntroductionTourSeen("Schematics"));
            return Task.CompletedTask;
        });
    }

    // The "Contribute" tab's tour on its own tab - its last step, about drafts, in the middle of the window.
    [Fact]
    public async Task The_contribute_tab_has_a_tour_of_its_own()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0");

        await this.WithMainAsync(main =>
        {
            OnTab(main, "ContributeTabItem");

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Assert.Equal("The Contribute tab", Overlay(main).GetControl<TextBlock>("HeadlineText").Text);

            ToFirstStep(main);
            Assert.Equal(TourCatalogue.Of(TourTab.Contribute).Steps.Count, StepCount(main));

            TourOverlay overlay = Overlay(main);

            while (overlay.StepShown < StepCount(main) - 1)
            {
                overlay.Next();
            }

            MainToursTests.Settle(main);
            Assert.Null(overlay.HoleShown);

            overlay.Next();
            Assert.Equal("4.0.0", UserSettings.IntroductionTourSeen("Contribute"));
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** ANOTHER TAB SHOWN UNDER A TOUR PUTS IT AWAY, REMEMBERING NOTHING (found in a render, 2026-11-11). *** Only code
    // can switch tabs while a tour shows - it takes every click - and the "Schematics" tour stood over the "Contribute"
    // tab, which then offered no tour of its own. It is offered again the next time its tab is shown.
    // ###########################################################################################
    [Fact]
    public async Task Another_tab_shown_under_a_tour_puts_it_away_remembering_nothing()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            ToFirstStep(main);

            OnTab(main, "ContributeTabItem");

            Assert.False(Overlay(main).IsShowing);
            Assert.False(Overlay(main).IsVisible);
            Assert.Null(UserSettings.IntroductionTourSeen("Schematics"));

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Assert.Equal("The Contribute tab", Overlay(main).GetControl<TextBlock>("HeadlineText").Text);
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** EVERY TAB HAS A TOUR, OR IS ONE THE OWNER LEFT WITHOUT (owner requests, 2026-11-11: "OK, now do the rest", and
    // 2026-11-12: "no need to have an introduction for the "About" and "Resources" tabs", the Maintainer tab "much more
    // limited and specialized"). *** A tab added to the window in neither list fails here.
    // ###########################################################################################
    [Fact]
    public void Every_tab_of_the_window_has_a_tour_or_is_left_without_on_purpose()
    {
        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            try
            {
                List<TabItem> tabs = main.GetControl<TabControl>("MainTabControl").Items.OfType<TabItem>().ToList();

                Assert.NotEmpty(tabs);
                Assert.Equal(Enum.GetValues<TourTab>().Length, CRT.Main.TourTabItems.Count);

                foreach (TabItem item in tabs)
                {
                    bool hasTour = CRT.Main.TourTabItems.Values.Any(name => ReferenceEquals(main.FindTourTarget(name), item));
                    bool leftWithout = CRT.Main.TabItemsWithoutTour.Any(name => ReferenceEquals(main.FindTourTarget(name), item));

                    Assert.True(hasTour != leftWithout, $"the tab [{item.Name}] must have a tour or be on the list of tabs without one - one of them");
                }
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // Another tab's tour shows on its own tab - "Feedback", which is always there.
    [Fact]
    public async Task Another_tab_shows_its_own_tour()
    {
        await this.WithMainAsync(main =>
        {
            OnTab(main, "FeedbackTabItem");

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            Assert.Equal("The Feedback tab", Overlay(main).GetControl<TextBlock>("HeadlineText").Text);
            return Task.CompletedTask;
        });
    }

    // The tabs left without a tour offer none, and remember nothing.
    [Theory]
    [InlineData("AboutTabItem")]
    [InlineData("ResourcesTabItem")]
    [InlineData("MaintainerTabItem")]
    public async Task A_tab_left_without_a_tour_shows_none(string tabItem)
    {
        UserSettings.EnableMaintainerTab = true;

        await this.WithMainAsync(main =>
        {
            OnTab(main, tabItem);

            Assert.False(main.OfferTourNow());
            Assert.False(Overlay(main).IsShowing);
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // *** THE FIRST INTRODUCTION TAKEN ENDS ON THE OTHER TABS (owner request, 2026-11-11: "As the last screen it should
    // show info that the user now can go to any other tab to see the introduction for that tab"). *** Its last card says
    // so and frames the other tabs; once another tab's tour has been seen, no tour ends on it.
    // ###########################################################################################
    [Fact]
    public async Task The_first_introduction_ends_on_the_other_tabs_and_no_later_one_does()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            TourOverlay overlay = Overlay(main);
            ToFirstStep(main);

            while (overlay.StepShown < StepCount(main) - 1)
            {
                overlay.Next();
            }

            MainToursTests.Settle(main);
            Assert.Equal(TourWording.OtherToursHeadline, overlay.GetControl<TextBlock>("HeadlineText").Text);
            Assert.Equal(TourWording.OtherToursBody, BoldRuns.TextOf(overlay.GetControl<TextBlock>("BodyText")));

            Rect hole = Assert.NotNull(overlay.HoleShown);
            Control contribute = main.GetControl<TabItem>("ContributeTabItem");
            Assert.True(hole.Contains(new Rect(contribute.TranslatePoint(new Point(0, 0), overlay)!.Value, contribute.Bounds.Size)), $"the hole {hole} does not hold the \"Contribute\" tab");

            overlay.Next();
            Assert.False(overlay.IsShowing);

            // The "Contribute" tab's tour, the "Schematics" one seen: its own steps, nothing more.
            OnTab(main, "ContributeTabItem");
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            ToFirstStep(main);
            Assert.Equal(TourCatalogue.Of(TourTab.Contribute).Steps.Count, StepCount(main));
            overlay.Skip();
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // A step about something further down a page that scrolls brings it into view, and scrolls the page back as it was once
    // left: the "Configuration" tab's buttons for the tours, the problems and the folders.
    // ###########################################################################################
    [Fact]
    public async Task A_step_further_down_a_page_scrolls_it_into_view_and_back()
    {
        await this.WithMainAsync(main =>
        {
            OnTab(main, "ConfigurationTabItem");
            ScrollViewer settings = main.TabConfiguration.GetControl<ScrollViewer>("SettingsScrollViewer");
            Vector before = settings.Offset;

            Assert.True(Offered(main), main.TourWaitReasonForTests());
            TourOverlay overlay = Overlay(main);
            int step = TourCatalogue.Of(TourTab.Configuration).Steps.ToList().FindIndex(each => each.Stage == TourStage.ScrolledIntoView);

            while (overlay.StepShown < step)
            {
                overlay.Next();
            }

            MainToursTests.Settle(main);
            Rect hole = Assert.NotNull(overlay.HoleShown);

            // The group whole: its first button and its last.
            foreach (string name in (string[])["TakeToursAgainButton", "OpenLogsFolderButton"])
            {
                Control button = main.TabConfiguration.GetControl<Button>(name);
                Assert.True(hole.Contains(new Rect(button.TranslatePoint(new Point(0, 0), overlay)!.Value, button.Bounds.Size)), $"the hole {hole} does not hold [{name}]");
            }

            Assert.True(settings.Offset.Y > before.Y, "the page was not scrolled");

            overlay.Skip();
            MainToursTests.Settle(main);
            Assert.Equal(before, settings.Offset);
            return Task.CompletedTask;
        });
    }

    // ###########################################################################################
    // "Take the introduction tours again" on the "Configuration" tab: every tour as never seen, and the "Schematics" tab
    // shown, where its tour comes again.
    // ###########################################################################################
    [Fact]
    public async Task Take_the_introduction_tours_again_forgets_them_and_shows_the_schematics_tab()
    {
        UserSettings.SetIntroductionTourSeen("Schematics", "4.0.0");
        UserSettings.SetIntroductionTourSeen("Contribute", "4.0.0");

        await this.WithMainAsync(main =>
        {
            OnTab(main, "ConfigurationTabItem");

            Button again = main.TabConfiguration.GetControl<Button>("TakeToursAgainButton");
            Assert.Equal(TourWording.TakeAgainButton, again.Content);
            again.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            MainToursTests.Settle(main);

            Assert.Null(UserSettings.IntroductionTourSeen("Schematics"));
            Assert.Null(UserSettings.IntroductionTourSeen("Contribute"));
            Assert.Same(main.GetControl<TabItem>("SchematicsTabItem"), main.GetControl<TabControl>("MainTabControl").SelectedItem);
            Assert.True(Offered(main), main.TourWaitReasonForTests());
            return Task.CompletedTask;
        });
    }

    // While a tour shows, the window's own keys do nothing - F11 would open the schematics' fullscreen under it.
    [Fact]
    public async Task While_a_tour_shows_the_windows_own_keys_do_nothing()
    {
        await this.WithMainAsync(main =>
        {
            Assert.True(Offered(main), main.TourWaitReasonForTests());

            main.KeyPress(Key.F11, RawInputModifiers.None, PhysicalKey.F11, keySymbol: null);
            main.KeyRelease(Key.F11, RawInputModifiers.None, PhysicalKey.F11, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            object? fullscreen = typeof(CRT.Main)
                .GetField("_schematicsFullscreenWindow", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .GetValue(main);

            Assert.Null(fullscreen);
            Assert.True(Overlay(main).IsShowing);
            return Task.CompletedTask;
        });
    }
}
