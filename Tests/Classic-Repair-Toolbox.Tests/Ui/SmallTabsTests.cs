using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using System.Reflection;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The four small tabs - Configuration, Overview, About and Resources - covering the parts that
// are real logic rather than launcher plumbing.
//
// These are the long tail of the coverage work, done last precisely because they are easy. What
// is worth testing here:
//
//  - Configuration is a settings surface. Its whole job is "seed the controls from UserSettings,
//    and write back when the user toggles something", so both directions are worth pinning: a
//    saved value must reach the control, and a click must reach the setting.
//  - Overview turns board data into printable rows and filters them.
//  - About renders the per-board credits list and revision date.
//
// What is deliberately NOT here: the Feedback tab (its substance is an HTTP upload and a
// file-picker dialog), Configuration's folder launchers and Overview's print path - all
// Process.Start or real network, which rule 6 in .claude/CLAUDE.md puts out of scope.
//
// COLLECTION NOTE: "HeadlessUi" because these construct controls. They also drive UserSettings'
// static state, which is safe only because xunit.runner.json disables collection parallelism;
// every test here points UserSettings at its own temp file first and restores nothing global.
// ###########################################################################################
[Collection("HeadlessUi")]
public class SmallTabsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SmallTabsTests()
    {
        this.RedirectSettingsToTemp();
    }

    public void Dispose()
    {
        this.RedirectSettingsToTemp();
        this.thisWorkspace.Dispose();
    }

    private void RedirectSettingsToTemp()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
    }

    // -----------------------------------------------------------------------------------------
    // Configuration - seeding the controls from saved settings
    // -----------------------------------------------------------------------------------------

    // The constructor deliberately assigns every checkbox BEFORE subscribing to its changed event,
    // so seeding does not fire the handlers and re-save what was just read. If that ordering is
    // ever reversed these still pass - but the pair of tests below (seed, then toggle) is what
    // proves both directions work.
    [Fact]
    public void Saved_settings_are_reflected_in_the_configuration_checkboxes()
    {
        UiTest.Run(() =>
        {
            UserSettings.CheckVersionOnLaunch = false;
            UserSettings.EnableWorklog = false;
            UserSettings.MultipleInstancesForComponentPopup = true;

            var tab = new TabConfiguration();

            Assert.False(tab.GetControl<CheckBox>("CheckVersionOnLaunchCheckBox").IsChecked);
            Assert.False(tab.GetControl<CheckBox>("EnableWorklogCheckBox").IsChecked);
            Assert.True(tab.GetControl<CheckBox>("MultipleInstancesForComponentPopupCheckBox").IsChecked);
        });
    }

    [Fact]
    public void The_theme_drop_down_is_seeded_from_the_saved_theme()
    {
        UiTest.Run(() =>
        {
            UserSettings.ThemeVariant = "Dark";

            var tab = new TabConfiguration();

            Assert.Equal(1, tab.GetControl<ComboBox>("ThemeVariantComboBox").SelectedIndex);
        });
    }

    // An unrecognised stored value must fall back to the first entry rather than leaving the box
    // unselected - a settings file can be hand-edited or carried over from an older version.
    [Fact]
    public void An_unknown_saved_theme_falls_back_to_the_first_entry()
    {
        UiTest.Run(() =>
        {
            UserSettings.ThemeVariant = "Chartreuse";

            var tab = new TabConfiguration();

            Assert.Equal(0, tab.GetControl<ComboBox>("ThemeVariantComboBox").SelectedIndex);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Configuration - writing settings back
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Ticking_a_configuration_checkbox_persists_the_setting()
    {
        UiTest.Run(() =>
        {
            UserSettings.MultipleInstancesForComponentPopup = false;

            var tab = new TabConfiguration();
            tab.GetControl<CheckBox>("MultipleInstancesForComponentPopupCheckBox").IsChecked = true;

            Assert.True(UserSettings.MultipleInstancesForComponentPopup);
        });
    }

    [Fact]
    public void Unticking_a_configuration_checkbox_persists_the_setting()
    {
        UiTest.Run(() =>
        {
            UserSettings.ShowDevelopmentVersionNotification = true;

            var tab = new TabConfiguration();
            tab.GetControl<CheckBox>("ShowDevelopmentVersionNotificationCheckBox").IsChecked = false;

            Assert.False(UserSettings.ShowDevelopmentVersionNotification);
        });
    }

    [Fact]
    public void Ticking_the_alpha_notification_checkbox_persists_the_setting()
    {
        UiTest.Run(() =>
        {
            UserSettings.AllowAlphaVersionNotification = false;

            var tab = new TabConfiguration();
            tab.GetControl<CheckBox>("AllowAlphaVersionNotificationCheckBox").IsChecked = true;

            Assert.True(UserSettings.AllowAlphaVersionNotification);
        });
    }

    [Fact]
    public void Choosing_a_theme_persists_it()
    {
        UiTest.Run(() =>
        {
            UserSettings.ThemeVariant = "Light";

            var tab = new TabConfiguration();
            tab.GetControl<ComboBox>("ThemeVariantComboBox").SelectedIndex = 1;

            Assert.Equal("Dark", UserSettings.ThemeVariant);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Configuration - dependent checkbox states
    // -----------------------------------------------------------------------------------------

    // Two settings only mean anything while launch-time data sync is on: the BETA source to sync
    // FROM, and the orphan cleanup that runs as PART of a sync. Both are disabled rather than
    // hidden when it is off, so the user can see they exist and why they are unavailable.
    [Fact]
    public void The_sync_dependent_checkboxes_are_disabled_when_launch_sync_is_off()
    {
        UiTest.Run(() =>
        {
            UserSettings.CheckDataOnLaunch = false;

            var tab = new TabConfiguration();

            Assert.False(tab.GetControl<CheckBox>("DownloadDataFromTestSourceCheckBox").IsEnabled);
            Assert.False(tab.GetControl<CheckBox>("AllowDeletionOfOrphanAndNonUsedFilesCheckBox").IsEnabled);
        });
    }

    [Fact]
    public void The_sync_dependent_checkboxes_are_enabled_when_launch_sync_is_on()
    {
        UiTest.Run(() =>
        {
            UserSettings.CheckDataOnLaunch = true;

            var tab = new TabConfiguration();

            Assert.True(tab.GetControl<CheckBox>("DownloadDataFromTestSourceCheckBox").IsEnabled);
            Assert.True(tab.GetControl<CheckBox>("AllowDeletionOfOrphanAndNonUsedFilesCheckBox").IsEnabled);
        });
    }

    // The public setter Main uses to push the value back when the sync banner turns it off.
    [Fact]
    public void Setting_the_data_on_launch_value_updates_the_checkbox()
    {
        UiTest.Run(() =>
        {
            UserSettings.CheckDataOnLaunch = true;
            var tab = new TabConfiguration();

            tab.SetCheckDataOnLaunchCheckBoxValue(false);

            Assert.False(tab.GetControl<CheckBox>("CheckDataOnLaunchCheckBox").IsChecked);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Overview
    // -----------------------------------------------------------------------------------------

    // One row per component, with its links gathered from the board's separate local-file and
    // link tables - the join that makes the printable list useful.
    [Fact]
    public void The_overview_builds_one_row_per_component_with_its_links()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();

            tab.LoadData(OverviewBoard());

            var rows = OverviewRows(tab);

            Assert.Equal(2, rows.Count);

            var u8 = rows.First(row => row.Component == "U8");
            Assert.Equal("SuperPLA", u8.FriendlyName);
            Assert.Equal(2, u8.Links.Count);

            // The component with no files or links of its own gets none of U8's.
            Assert.Empty(rows.First(row => row.Component == "C1").Links);
        });
    }

    // Rows start ticked for printing, so "print the list" works without the user selecting
    // anything first.
    [Fact]
    public void Overview_rows_start_selected_for_print()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();

            tab.LoadData(OverviewBoard());

            Assert.All(OverviewRows(tab), row => Assert.True(row.IsSelectedForPrint));
        });
    }

    // Search terms are ANDed across the row's whole display string, so a term matching the
    // friendly name and one matching the part number together still find the row.
    [Fact]
    public void The_overview_search_narrows_to_matching_rows()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();
            tab.LoadData(OverviewBoard());

            tab.ApplyFilter("SuperPLA");

            var rows = OverviewRows(tab);
            Assert.Single(rows);
            Assert.Equal("U8", rows[0].Component);
        });
    }

    [Fact]
    public void The_overview_search_is_case_insensitive()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();
            tab.LoadData(OverviewBoard());

            tab.ApplyFilter("superpla");

            Assert.Single(OverviewRows(tab));
        });
    }

    // Several terms must ALL match, and a term matching nothing empties the list rather than
    // falling back to everything.
    [Fact]
    public void Overview_search_terms_are_all_required()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();
            tab.LoadData(OverviewBoard());

            tab.ApplyFilter("SuperPLA nonsense");

            Assert.Empty(OverviewRows(tab));
        });
    }

    // An empty box is not a filter.
    [Fact]
    public void An_empty_overview_search_shows_every_row()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();
            tab.LoadData(OverviewBoard());

            tab.ApplyFilter("   ");

            Assert.Equal(2, OverviewRows(tab).Count);
        });
    }

    // Loading a different board replaces the previous board's rows rather than appending to them.
    [Fact]
    public void Loading_a_second_board_replaces_the_overview_rows()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOverview();
            tab.LoadData(OverviewBoard());

            var otherBoard = new BoardData();
            otherBoard.Components.Add(new ComponentEntry { BoardLabel = "R1", FriendlyName = "Resistor" });
            tab.LoadData(otherBoard);

            var rows = OverviewRows(tab);
            Assert.Single(rows);
            Assert.Equal("R1", rows[0].Component);
        });
    }

    // -----------------------------------------------------------------------------------------
    // About
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_about_tab_shows_the_supplied_version()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.InitializeAbout(Assembly.GetExecutingAssembly(), "2.5.0");

            Assert.Equal("2.5.0", tab.GetControl<TextBlock>("AppVersionText").Text);
        });
    }

    // A null version must not render as an empty gap - the About tab is where a user is asked to
    // report their version from.
    [Fact]
    public void A_missing_version_is_shown_as_unknown_rather_than_blank()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.InitializeAbout(Assembly.GetExecutingAssembly(), null);

            Assert.Equal("(unknown)", tab.GetControl<TextBlock>("AppVersionText").Text);
        });
    }

    [Fact]
    public void A_board_revision_date_is_shown_when_present()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.SetBoardInfo("2026-01-15", credits: null);

            Assert.True(tab.GetControl<Control>("RevisionDatePanel").IsVisible);
            Assert.Equal("2026-01-15", tab.GetControl<TextBlock>("RevisionDateText").Text);
        });
    }

    // Most boards carry no revision date, so the row collapses rather than showing a label with
    // nothing after it.
    [Fact]
    public void The_revision_date_row_is_hidden_when_there_is_no_date()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.SetBoardInfo("   ", credits: null);

            Assert.False(tab.GetControl<Control>("RevisionDatePanel").IsVisible);
        });
    }

    [Fact]
    public void Board_credits_are_listed()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.SetBoardInfo("2026-01-15", new List<CreditEntry>
            {
                new() { Category = "Schematics", NameOrHandle = "Someone", Contact = "https://example.com" },
                new() { Category = "Photos", NameOrHandle = "Another", Contact = string.Empty },
            });

            Assert.Equal(2, tab.CreditsList.Count);
            Assert.True(tab.GetControl<Control>("CreditsSectionBorder").IsVisible);
        });
    }

    // A contact that is a web address or an email is clickable; anything else is plain text
    // rather than a dead link.
    //
    // Note what counts as an email: ContactLinkFormatter.IsContactEmail is "contains @ and no
    // space", so a bare handle like "@someone" IS treated as one and becomes a mailto:. That is
    // deliberate and documented there - the no-space rule exists to stop a sentence that merely
    // mentions an address being linkified, not to validate the address. This test pins the real
    // rule down rather than an idealised one; a stricter check would be a behaviour change.
    [Fact]
    public void Only_a_web_or_email_contact_is_clickable()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.SetBoardInfo(null, new List<CreditEntry>
            {
                new() { Category = "A", NameOrHandle = "Web", Contact = "https://example.com" },
                new() { Category = "B", NameOrHandle = "Mail", Contact = "someone@example.com" },
                new() { Category = "C", NameOrHandle = "Handle", Contact = "@someone_on_a_forum" },
                new() { Category = "D", NameOrHandle = "Prose", Contact = "ask on the forum" },
                new() { Category = "E", NameOrHandle = "None", Contact = string.Empty },
            });

            Assert.True(tab.CreditsList[0].IsLink);
            Assert.True(tab.CreditsList[1].IsLink);

            // Contains @ and no space, so it is treated as an email - see the comment above.
            Assert.True(tab.CreditsList[2].IsLink);

            Assert.False(tab.CreditsList[3].IsLink);
            Assert.False(tab.CreditsList[4].IsLink);
        });
    }

    // With no credits the whole section is hidden - a board with none should not show an empty
    // "Credits" heading.
    [Fact]
    public void The_credits_section_is_hidden_when_the_board_has_none()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.SetBoardInfo("2026-01-15", new List<CreditEntry>());

            Assert.False(tab.GetControl<Control>("CreditsSectionBorder").IsVisible);
        });
    }

    // Switching board replaces the previous board's credits rather than accumulating them.
    [Fact]
    public void Loading_a_second_boards_credits_replaces_the_first()
    {
        UiTest.Run(() =>
        {
            var tab = new TabAbout();

            tab.SetBoardInfo(null, new List<CreditEntry>
            {
                new() { Category = "A", NameOrHandle = "First" },
            });

            tab.SetBoardInfo(null, new List<CreditEntry>
            {
                new() { Category = "B", NameOrHandle = "Second" },
            });

            Assert.Single(tab.CreditsList);
            Assert.Equal("Second", tab.CreditsList[0].Name);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Fixture
    // -----------------------------------------------------------------------------------------

    // A board with two components, one of which carries a local file and a web link so the
    // link-gathering join has something to find - and something to leave alone.
    private static BoardData OverviewBoard()
    {
        var board = new BoardData();

        board.Components.Add(new ComponentEntry
        {
            BoardLabel = "U8",
            FriendlyName = "SuperPLA",
            TechnicalNameOrValue = "251715",
            Category = "IC",
            PartNumber = "906114-01",
            Description = "Programmable logic array",
        });

        board.Components.Add(new ComponentEntry
        {
            BoardLabel = "C1",
            FriendlyName = "Capacitor",
            TechnicalNameOrValue = "100nF",
            Category = "Capacitor",
        });

        board.ComponentLocalFiles.Add(new ComponentLocalFileEntry
        {
            BoardLabel = "U8",
            Name = "Datasheet",
            File = "docs/pla.pdf",
        });

        board.ComponentLinks.Add(new ComponentLinkEntry
        {
            BoardLabel = "U8",
            Name = "Pinout",
            Url = "https://example.com/pla",
        });

        return board;
    }

    private static List<OverviewRow> OverviewRows(TabOverview tab)
    {
        var itemsSource = tab.GetControl<ItemsControl>("OverviewItemsControl").ItemsSource;

        return itemsSource is null
            ? new List<OverviewRow>()
            : itemsSource.Cast<OverviewRow>().ToList();
    }
}
