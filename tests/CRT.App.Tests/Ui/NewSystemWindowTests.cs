using Avalonia.Controls;
using Avalonia.Headless;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// NewSystemWindow - the "Add a new system" dialog (NewContributeStrategy.md Phase 2, session 2c,
// task 9). Drives the real validation path: the Create button's enabled state and the registration
// the window would hand back are both produced by the same BuildRegistrationOrNull the click uses,
// so an input the button accepts is by construction one the create path can act on.
//
// The three names become real FOLDER segments, so the rejection cases here are not cosmetic - each
// one prevents a system that looks created and is then broken in a way the contributor cannot
// diagnose. NewSystemIdentityTests covers the rules themselves; these cover the window applying
// them, disabling the button and saying why.
//
// Drives DraftManager's real static state via LoadFrom (its own seam) because the window checks
// whether a draft folder is already occupied.
//
// *** KNOWN INTERMITTENT: A_reserved_device_name_is_rejected. *** Observed failing twice in about
// fifteen full-solution runs on 2026-09-21, and NOT reproducible on demand - it passes alone,
// passes with this class run repeatedly, and passes on three consecutive full CRT.App runs. It is
// recorded here rather than left to be rediscovered because a flake nobody has named gets
// re-investigated from scratch every time.
//
// The leading suspicion, unproven: ManufacturerBox is an AutoCompleteBox carrying an ItemsSource
// with MinimumPrefixLength="0", so it re-filters on EVERY text change including an empty one, and
// Avalonia runs that filtering on a delayed dispatcher pass which can write back to .Text. Fill()
// sets .Text and then calls RefreshValidationForTests synchronously; if a pending filter pass from
// a PREVIOUS test in the shared "HeadlessUi" dispatcher lands in between, validation could read a
// value the test did not set. That would explain why it needs the full suite to appear.
//
// Whoever picks this up: the cheapest probe is to assert ManufacturerBox.Text immediately before
// the validation call inside the failing test. If it is not "CON", the theory holds and the fix
// belongs in Fill() (pump the dispatcher after setting text), not in the window.
[Collection("HeadlessUi")]
public sealed class NewSystemWindowTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public NewSystemWindowTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string hardware, string board, string excelDataFile) => new()
    {
        HardwareName = hardware,
        BoardName = board,
        ExcelDataFile = excelDataFile,
    };

    private static NewSystemWindow WindowWith(params HardwareBoardEntry[] existing)
    {
        var window = new NewSystemWindow();
        window.Initialize(existing);
        return window;
    }

    // A window that is never attached to a visual tree never raises TextChanged, so setting .Text
    // here does not reach the window's own handler - the same trap TabWorkbooks documents for its
    // search box. RefreshValidationForTests is what that handler calls, so this still drives the
    // shipped validation path rather than a parallel copy of it.
    private static void Fill(NewSystemWindow window, string manufacturer, string hardware, string board)
    {
        window.GetControl<AutoCompleteBox>("ManufacturerBox").Text = manufacturer;
        window.GetControl<TextBox>("HardwareBox").Text = hardware;
        window.GetControl<TextBox>("BoardBox").Text = board;
        window.RefreshValidationForTests();
    }

    private static bool CreateEnabled(NewSystemWindow window) =>
        window.GetControl<Button>("CreateButton").IsEnabled;

    // ------------------------------------------------------------------ Opening state

    [Fact]
    public void The_window_constructs_without_throwing()
    {
        UiTest.Run(() => Assert.NotNull(WindowWith()));
    }

    // Nothing has been typed yet, so there is nothing to create - and deliberately nothing to
    // complain about either: an error message on an untouched form reads as a fault.
    [Fact]
    public void It_opens_with_create_disabled_and_no_complaint()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Assert.False(CreateEnabled(window));
            Assert.Equal(string.Empty, window.ValidationTextForTests);
        });
    }

    // ------------------------------------------------------------------ The happy path

    [Fact]
    public void Filling_in_all_three_names_enables_create()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commodore", "C64", "250407");

            Assert.True(CreateEnabled(window));
            Assert.Equal(string.Empty, window.ValidationTextForTests);
        });
    }

    [Fact]
    public void Leaving_any_one_name_blank_keeps_create_disabled()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commodore", "C64", "   ");
            Assert.False(CreateEnabled(window));

            Fill(window, "Commodore", string.Empty, "250407");
            Assert.False(CreateEnabled(window));

            Fill(window, string.Empty, "C64", "250407");
            Assert.False(CreateEnabled(window));
        });
    }

    // The preview is what makes the folder layout not a surprise - the contributor sees the exact
    // identity that is about to be created, including a typo, before it becomes a folder on disk.
    [Fact]
    public void The_preview_shows_the_identity_that_will_be_created()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commodore", "C64", "250407");

            Assert.Contains("Commodore/C64/250407/Data C64 250407.xlsx", window.PreviewTextForTests);
        });
    }

    [Fact]
    public void The_registration_the_window_would_return_carries_the_typed_values()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();
            window.GetControl<TextBox>("NotesBox").Text = "My own notes";
            Fill(window, "Commodore", "C64", "250407");

            var registration = window.BuildRegistrationOrNull(out string message);

            Assert.NotNull(registration);
            Assert.Equal(string.Empty, message);
            Assert.Equal("C64", registration!.HardwareName);
            Assert.Equal("250407", registration.BoardName);
            Assert.Equal("My own notes", registration.HardwareNotes);
            Assert.Equal("Commodore/C64/250407/Data C64 250407.xlsx", registration.ExcelDataFile);
            Assert.NotEqual(string.Empty, registration.CreatedUtc);
        });
    }

    // ------------------------------------------------------------------ Rejections

    // One check covers all three sources of an existing name: the main workbook, a legacy
    // _UserContribution sidecar, and another draft's own registration - by the time this window
    // opens, DataManager.HardwareBoards already holds all three.
    [Fact]
    public void A_hardware_and_board_that_already_exists_is_rejected_by_name()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith(Entry("C64", "250407", "Commodore/C64/250407/Data.xlsx"));

            Fill(window, "Commodore", "C64", "250407");

            Assert.False(CreateEnabled(window));
            Assert.Contains("already exists", window.ValidationTextForTests);
            Assert.Contains("C64", window.ValidationTextForTests);
            Assert.Contains("250407", window.ValidationTextForTests);
        });
    }

    [Fact]
    public void The_duplicate_check_ignores_casing()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith(Entry("C64", "250407", "Commodore/C64/250407/Data.xlsx"));

            Fill(window, "Commodore", "c64", "250407");

            Assert.False(CreateEnabled(window));
        });
    }

    // A different board on the same hardware is fine - that is the ordinary case for someone
    // documenting a second revision of a machine already in the list.
    [Fact]
    public void A_new_board_on_existing_hardware_is_accepted()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith(Entry("C64", "250407", "Commodore/C64/250407/Data.xlsx"));

            Fill(window, "Commodore", "C64", "250425");

            Assert.True(CreateEnabled(window));
        });
    }

    [Theory]
    [InlineData("Commodore", "C64/NTSC", "250407")]
    [InlineData("Commodore", "C64", "250407:rev")]
    [InlineData("Commo*dore", "C64", "250407")]
    public void A_name_with_a_path_breaking_character_is_rejected_and_says_so(
        string manufacturer, string hardware, string board)
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, manufacturer, hardware, board);

            Assert.False(CreateEnabled(window));
            Assert.Contains("cannot contain", window.ValidationTextForTests);
        });
    }

    // Windows silently strips a trailing dot, so the folder created would no longer match the key
    // recorded in draft.json and the draft would become unfindable by its own lookup.
    [Fact]
    public void A_name_ending_in_a_dot_is_rejected()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commodore", "C64.", "250407");

            Assert.False(CreateEnabled(window));
            Assert.Contains("dot", window.ValidationTextForTests);
        });
    }

    [Fact]
    public void A_reserved_device_name_is_rejected()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commodore", "CON", "250407");

            Assert.False(CreateEnabled(window));
            Assert.Contains("reserved", window.ValidationTextForTests);
        });
    }

    // The message names WHICH field is wrong - with three free-text boxes, "invalid name" alone
    // would leave the contributor hunting.
    [Fact]
    public void The_message_names_the_field_that_is_wrong()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commo|dore", "C64", "250407");
            Assert.StartsWith("Manufacturer", window.ValidationTextForTests);

            Fill(window, "Commodore", "C6|4", "250407");
            Assert.StartsWith("Hardware", window.ValidationTextForTests);

            Fill(window, "Commodore", "C64", "2504|07");
            Assert.StartsWith("Board", window.ValidationTextForTests);
        });
    }

    // A folder that already holds a draft belongs to a live system - creating over it would
    // overwrite that draft's own draft.json.
    [Fact]
    public void A_folder_that_already_holds_a_draft_is_rejected()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = NewSystemIdentity.BuildExcelDataFile("Commodore", "C64", "250407");
            // The MARKER is what makes a folder a draft since Phase 6 - a stray workbook copied
            // in by hand is not one, and must not block a legitimate name.
            DraftMarkerStore.Save(
                DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
                new DraftMarker { SystemKey = excelDataFile });

            var window = WindowWith();
            Fill(window, "Commodore", "C64", "250407");

            Assert.False(CreateEnabled(window));
            Assert.Contains("already exists in that folder", window.ValidationTextForTests);
        });
    }

    // Correcting the mistake has to clear the complaint and re-enable the button, or the form
    // would look permanently stuck after one bad keystroke.
    [Fact]
    public void Correcting_a_rejected_name_enables_create_again()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith();

            Fill(window, "Commodore", "C64/NTSC", "250407");
            Assert.False(CreateEnabled(window));

            Fill(window, "Commodore", "C64", "250407");

            Assert.True(CreateEnabled(window));
            Assert.Equal(string.Empty, window.ValidationTextForTests);
        });
    }

    // ------------------------------------------------------------------ Manufacturer suggestions

    // Offering the manufacturers already in use is what stops a typo quietly creating a second
    // top-level folder beside the real one - nothing else would catch "Comodore".
    [Fact]
    public void Existing_manufacturers_are_offered_as_suggestions_without_duplicates()
    {
        UiTest.Run(() =>
        {
            var window = WindowWith(
                Entry("C64", "250407", "Commodore/C64/250407/Data.xlsx"),
                Entry("C64", "250425", "Commodore/C64/250425/Data.xlsx"),
                Entry("CPC464", "MC0001A", "Amstrad/CPC464/MC0001A/Data.xlsx"));

            var suggestions = window.GetControl<AutoCompleteBox>("ManufacturerBox")
                .ItemsSource!
                .Cast<string>()
                .ToList();

            Assert.Equal(new[] { "Amstrad", "Commodore" }, suggestions);
        });
    }

    // ------------------------------------------------------------------ The maintainer agreement
    //
    // *** A NEW SYSTEM CAN ONLY BE CREATED BY ACCEPTING THE MAINTAINER ROLE (owner request,
    // 2026-09-24). *** "Create system" shows NewSystemMaintainerWindow, and the window hands back a
    // registration - the only thing Main creates a system from - only when that is accepted.
    //
    // Most of these replace the agreement with a canned answer (ConfirmMaintainerRoleOverrideForTests)
    // and drive the real ConfirmAndCloseAsync around it. The last two do NOT: they go through the
    // real agreement dialog, which is what proves the default path actually shows it rather than
    // the seam being the only thing that ever asks.

    [Fact]
    public async Task Declining_the_agreement_creates_nothing_and_keeps_the_form_filled_in()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = WindowWith();
            window.Show();
            window.ConfirmMaintainerRoleOverrideForTests = _ => Task.FromResult(false);
            Fill(window, "Commodore", "C64", "250407");

            NewSystemRegistration? result = await window.ConfirmAndCloseAsync();

            Assert.Null(result);

            // Back on the form, with everything still typed in - declining is not a reason to
            // make the contributor start over.
            Assert.True(window.IsVisible, "Declining must leave the form open.");
            Assert.Equal("C64", window.GetControl<TextBox>("HardwareBox").Text);
            Assert.True(CreateEnabled(window));

            window.Close();
        });
    }

    [Fact]
    public async Task Accepting_the_agreement_hands_back_the_registration_and_closes()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = WindowWith();
            window.Show();
            window.ConfirmMaintainerRoleOverrideForTests = _ => Task.FromResult(true);
            Fill(window, "Commodore", "C64", "250407");

            NewSystemRegistration? result = await window.ConfirmAndCloseAsync();

            Assert.NotNull(result);
            Assert.Equal("Commodore/C64/250407/Data C64 250407.xlsx", result!.ExcelDataFile);
            Assert.False(window.IsVisible);
        });
    }

    // The agreement is about ONE system, so it must be asked about the one actually being created.
    [Fact]
    public async Task The_agreement_is_asked_about_the_system_being_created()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = WindowWith();
            window.Show();

            NewSystemRegistration? askedAbout = null;
            window.ConfirmMaintainerRoleOverrideForTests = registration =>
            {
                askedAbout = registration;
                return Task.FromResult(false);
            };

            Fill(window, "Commodore", "C64", "250407");
            await window.ConfirmAndCloseAsync();

            Assert.NotNull(askedAbout);
            Assert.Equal("C64", askedAbout!.HardwareName);
            Assert.Equal("250407", askedAbout.BoardName);

            window.Close();
        });
    }

    // Invalid input is refused before the agreement - asking someone to accept a role for a system
    // that cannot be created would be asking for nothing.
    [Fact]
    public async Task Invalid_input_never_reaches_the_agreement()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = WindowWith();

            int asked = 0;
            window.ConfirmMaintainerRoleOverrideForTests = _ =>
            {
                asked++;
                return Task.FromResult(true);
            };

            Fill(window, "Commodore", "C64", "   ");

            Assert.Null(await window.ConfirmAndCloseAsync());
            Assert.Equal(0, asked);
        });
    }

    // ###########################################################################################
    // Through the REAL agreement dialog: Create opens NewSystemMaintainerWindow over the form,
    // naming the system, and Escape there declines - leaving the form open with nothing created.
    // ###########################################################################################
    [Fact]
    public async Task Create_really_shows_the_agreement_and_Escape_there_declines()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = WindowWith();
            window.Show();
            Fill(window, "Commodore", "C64", "250407");

            Task<NewSystemRegistration?> creating = window.ConfirmAndCloseAsync();

            NewSystemMaintainerWindow agreement = Assert.Single(window.OwnedWindows.OfType<NewSystemMaintainerWindow>());
            Assert.Equal("C64 - 250407", agreement.GetControl<TextBlock>("SystemNameText").Text);

            agreement.KeyPress(Avalonia.Input.Key.Escape, Avalonia.Input.RawInputModifiers.None, Avalonia.Input.PhysicalKey.Escape, keySymbol: null);

            Assert.Null(await creating);
            Assert.True(window.IsVisible, "Declining must leave the form open.");

            window.Close();
        });
    }

    [Fact]
    public async Task Create_through_the_real_agreement_completes_when_Accept_is_clicked()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = WindowWith();
            window.Show();
            Fill(window, "Commodore", "C64", "250407");

            Task<NewSystemRegistration?> creating = window.ConfirmAndCloseAsync();

            NewSystemMaintainerWindow agreement = Assert.Single(window.OwnedWindows.OfType<NewSystemMaintainerWindow>());
            agreement.GetControl<Button>("AcceptButton")
                .RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            NewSystemRegistration? result = await creating;

            Assert.NotNull(result);
            Assert.Equal("C64", result!.HardwareName);
            Assert.False(window.IsVisible);
        });
    }
}
