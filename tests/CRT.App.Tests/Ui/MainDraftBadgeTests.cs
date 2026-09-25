using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Draft" chip on the Hardware and Board drop-downs (owner request, 2026-09-24), on the
// real main window. DraftBadgeSetTests covers WHICH names are badged; these cover the drop-downs
// actually drawing it, and - the part that is easy to get wrong - drawing it on a drop-down that
// was ALREADY RENDERED when the draft appeared.
//
// Built the way MainWindowTests builds Main: constructed, never StartAsync'd (that reaches the
// network), with worklog and settings redirected to the workspace. The drafts are real markers
// on disk, paired with a HardwareBoardsOverrideForTests list, which is how the Drafts tab - and
// so the chips, which follow its list - see them.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainDraftBadgeTests : IDisposable
{
    // Names no other test uses, so nothing left in DataManager's static board list by another
    // class can match them and start a board load.
    private const string DraftedHardware = "Badge Test HW";
    private const string DraftedBoard = "Badge Board 1";
    private const string OtherHardware = "Badge Test Other HW";

    private readonly TempWorkspace thisWorkspace = new();

    private static string SystemKey => $"Test Manu/{DraftedHardware}/{DraftedBoard}/Data {DraftedHardware} {DraftedBoard}.xlsx";

    public MainDraftBadgeTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.Path_("settings.json"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.Path_("settings-after.json"));
        this.thisWorkspace.Dispose();
    }

    private static CRT.Main BuildMain()
    {
        var window = new CRT.Main();

        window.TabDrafts.HardwareBoardsOverrideForTests =
        [
            new HardwareBoardEntry
            {
                HardwareName = DraftedHardware,
                BoardName = DraftedBoard,
                ExcelDataFile = SystemKey,
            },
        ];

        return window;
    }

    private static void WriteDraftMarker() =>
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, SystemKey),
            new DraftMarker
            {
                SystemKey = SystemKey,
                BaseRevision = "2026-09-01",
                CreatedUtc = "2026-09-24T00:00:00Z",
            });

    private static bool HasChip(Control control) =>
        control.GetVisualDescendants()
            .OfType<Border>()
            .Any(border => border.Classes.Contains("DraftChip") && border.IsVisible);

    private static bool TemplateDrawsChip(ComboBox comboBox, string name) =>
        comboBox.ItemTemplate!.Build(name) is Panel panel &&
        panel.Children.OfType<Border>().Any(border => border.Classes.Contains("DraftChip"));

    // ###########################################################################################
    // The template itself: the drafted hardware and its drafted board get the chip, the others
    // do not. Read straight off each drop-down's own ItemTemplate - what it hands the drop-down
    // list and the closed box alike.
    // ###########################################################################################
    [Fact]
    public void A_drafted_hardware_and_board_are_drawn_with_the_chip_and_others_without()
    {
        WriteDraftMarker();

        UiTest.Run(() =>
        {
            CRT.Main window = BuildMain();

            window.HardwareComboBox.ItemsSource = new[] { DraftedHardware, OtherHardware };
            window.HardwareComboBox.SelectedIndex = 0;

            window.ApplyDraftsTabVisibility();

            Assert.True(TemplateDrawsChip(window.HardwareComboBox, DraftedHardware));
            Assert.False(TemplateDrawsChip(window.HardwareComboBox, OtherHardware));

            Assert.True(TemplateDrawsChip(window.BoardComboBox, DraftedBoard));
            Assert.False(TemplateDrawsChip(window.BoardComboBox, "Some other board"));
        });
    }

    // ###########################################################################################
    // *** THE REFRESH. *** The drop-down is already on screen, showing the selected hardware with
    // no chip; a draft then appears (the first "Save to draft" on a board does exactly this). The
    // chip must appear in the CLOSED box without the selection changing - re-assigning the items
    // instead would fire SelectionChanged and reload the board.
    // ###########################################################################################
    [Fact]
    public void A_draft_appearing_later_puts_the_chip_on_the_already_rendered_selection()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = BuildMain();
            window.Show();

            try
            {
                window.HardwareComboBox.ItemsSource = new[] { DraftedHardware, OtherHardware };
                window.HardwareComboBox.SelectedIndex = 0;

                window.ApplyDraftsTabVisibility();
                Dispatcher.UIThread.RunJobs();

                Assert.False(HasChip(window.HardwareComboBox), "No draft exists yet, so no chip.");

                int selectionChanges = 0;
                window.HardwareComboBox.SelectionChanged += (_, _) => selectionChanges++;

                WriteDraftMarker();
                window.ApplyDraftsTabVisibility();
                Dispatcher.UIThread.RunJobs();

                Assert.True(HasChip(window.HardwareComboBox), "The chip should appear on the closed drop-down.");
                Assert.Equal(0, selectionChanges);
                Assert.Equal(DraftedHardware, window.HardwareComboBox.SelectedItem);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // And the reverse: discarding the last draft takes the chip away again.
    [Fact]
    public void Discarding_the_draft_takes_the_chip_away_again()
    {
        WriteDraftMarker();

        UiTest.Run(() =>
        {
            CRT.Main window = BuildMain();
            window.Show();

            try
            {
                window.HardwareComboBox.ItemsSource = new[] { DraftedHardware, OtherHardware };
                window.HardwareComboBox.SelectedIndex = 0;

                window.ApplyDraftsTabVisibility();
                Dispatcher.UIThread.RunJobs();
                Assert.True(HasChip(window.HardwareComboBox));

                DraftManager.DiscardDraft(SystemKey);
                window.ApplyDraftsTabVisibility();
                Dispatcher.UIThread.RunJobs();

                Assert.False(HasChip(window.HardwareComboBox));
            }
            finally
            {
                window.Close();
            }
        });
    }
}
