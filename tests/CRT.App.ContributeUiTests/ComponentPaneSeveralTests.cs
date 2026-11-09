using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Avalonia;
using Avalonia.Headless;
using Avalonia.Input;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// SEVERAL COMPONENTS IN THE COMPONENT PANE, EDITED AS ONE (ComponentPane.Several.cs; owner request, 2026-11-08: "selecting
// components "R21" and "R22" and I go to "Basics" in UI view, then it should disable the "Board label" ... if I e.g. change
// "Category" or "Friendly name" fields, then both components will get this new value ... If I go to "Replacement parts",
// then I think it should remove all replacements line which are not part of all of them ... It must be very clear for
// the user, that whatever the user changes will affect all selected components").
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentPaneSeveralTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private static readonly string[] Both = ["R21", "R22"];

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Published() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "R21", Category = "Resistor", FriendlyName = "Pull-up", TechnicalNameOrValue = "4K7", Package = "Axial" },
            new ComponentEntry { BoardLabel = "R22", Category = "Resistor", FriendlyName = "Pull-down", TechnicalNameOrValue = "4K7", Package = "Axial" },
            new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Package = "DIP-40" },
        ],
        ComponentReplacements =
        [
            new ComponentReplacementEntry { BoardLabel = "R21", PartNumber = "MF 1%", Kind = "Drop-in replacement" },
            new ComponentReplacementEntry { BoardLabel = "R21", PartNumber = "Only R21" },
            new ComponentReplacementEntry { BoardLabel = "R22", PartNumber = "MF 1%", Kind = "Drop-in replacement" },
        ],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "R21", Pin = "1", Name = "Bus", File = "Scope baseline/r21.png" },
        ],
    };

    private DraftSession OpenDraft()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, Published());

        return DraftSession.Open(this.DraftsRoot, BoardKey)!;
    }

    private static (ComponentPane Pane, Window Window) Shown(DraftSession session, params string[] labels)
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.ShowSeveral(session, labels.Length > 0 ? labels : Both, string.Empty, Published());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    private static T Get<T>(ComponentPane pane, string name) where T : Control => pane.GetControl<T>(name);

    private static ComponentEntry Row(DraftSession session, string label) => ComponentPaneEdits.RowShown(session.Content.Board, label, string.Empty)!;

    private static List<string> PartsOf(DraftSession session, string label) =>
        ComponentPaneEdits.ReplacementsOf(session.Content.Board, label, string.Empty).Select(row => row.PartNumber).ToList();

    private static List<ComponentPaneReplacementRow> ReplacementRows(ComponentPane pane) =>
        Get<ItemsControl>(pane, "ReplacementsItemsControl").Items.Cast<ComponentPaneReplacementRow>().ToList();

    private static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // It is clear who is edited and that a change reaches them all; the label is off; nothing acts on one of them.
    [Fact]
    public void Several_shown_say_who_they_are_and_that_every_change_reaches_all_of_them()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.True(pane.IsSeveral);
            Assert.Equal(Both, pane.Labels);
            Assert.Equal("Components: R21, R22", Get<TextBlock>(pane, "TitleText").Text);
            Assert.True(Get<Border>(pane, "SeveralBar").IsVisible);

            // How many and who in bold - the number and each name, nothing between them (owner requests, 2026-11-09).
            TextBlock several = Get<TextBlock>(pane, "SeveralText");
            Assert.Contains("every change you make here is made to all of them", BoldRuns.TextOf(several), StringComparison.Ordinal);
            Assert.Equal(["2", "R21", "R22"], BoldRuns.BoldPiecesOf(several));

            TextBox label = Get<TextBox>(pane, "BoardLabelBox");
            Assert.False(label.IsEnabled);
            Assert.Equal("R21, R22", label.Text);

            Assert.False(Get<WrapPanel>(pane, "ActionsPanel").IsVisible);
            Assert.False(Get<Button>(pane, "DeleteComponentButton").IsVisible);

            // Shown again, the same components keep the very list "Make all like" picks from - refilled, each refill posted
            // a commit of every list of every one of them (code review, 2026-11-09).
            object? listed = Get<ComboBox>(pane, "MakeAllLikeSourceBox").ItemsSource;
            pane.ShowSeveral(session, Both, string.Empty, Published());
            Dispatcher.UIThread.RunJobs();
            Assert.Same(listed, Get<ComboBox>(pane, "MakeAllLikeSourceBox").ItemsSource);

            window.Close();
        });
    }

    // What they share shows; what differs is empty and says so.
    [Fact]
    public void A_field_they_share_shows_and_one_they_differ_in_says_so()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal("4K7", Get<TextBox>(pane, "TechnicalNameBox").Text);
            Assert.Equal("Resistor", Get<ComboBox>(pane, "CategoryBox").SelectedItem);

            TextBox friendly = Get<TextBox>(pane, "FriendlyNameBox");
            Assert.Equal(string.Empty, friendly.Text);
            Assert.Equal("Differs - 2 values", friendly.PlaceholderText);
            Assert.Null(ToolTip.GetTip(friendly));

            // The box says it; the bar does not count them (owner request, 2026-11-09: "remove the text "1 field here
            // differs between..." - no need for that").
            Assert.False(Get<TextBlock>(pane, "SeveralPageText").IsVisible);

            // Left as it is, a field that differs sets nothing.
            pane.CommitHeld();
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    [Fact]
    public void A_field_typed_is_set_on_all_of_them_as_one_step()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            TextBox friendly = Get<TextBox>(pane, "FriendlyNameBox");
            friendly.Text = "Bus resistor";
            pane.CommitHeld();

            Assert.Equal("Bus resistor", Row(session, "R21").FriendlyName);
            Assert.Equal("Bus resistor", Row(session, "R22").FriendlyName);
            Assert.Equal("CPU", Row(session, "U8").FriendlyName);
            Assert.Equal("Friendly name of [R21, R22]", session.UndoDescription);

            // Alike now: the box is a plain box again.
            Assert.NotEqual("Differs - 2 values", friendly.PlaceholderText);

            Assert.True(session.Undo());
            Assert.False(session.CanUndo);
            Assert.Equal("Pull-up", Row(session, "R21").FriendlyName);

            window.Close();
        });
    }

    // Only the rows all of them have are listed - and a change, a removal or an addition there reaches every one.
    [Fact]
    public void Replacement_parts_list_what_all_have_and_a_change_there_reaches_every_one()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.ShowPage(ComponentPage.Replacements);

            ComponentPaneReplacementRow common = Assert.Single(ReplacementRows(pane));
            Assert.Equal("MF 1%", common.PartNumber);
            Assert.Contains("R21: Only R21", Get<TextBlock>(pane, "SeveralPageText").Text, StringComparison.Ordinal);

            common.PartNumber = "MF 0.5%";
            pane.CommitHeld();

            Assert.Equal(["MF 0.5%", "Only R21"], PartsOf(session, "R21"));
            Assert.Equal(["MF 0.5%"], PartsOf(session, "R22"));

            // Added: to every one.
            Click(Get<Button>(pane, "AddReplacementButton"));
            ReplacementRows(pane)[^1].PartNumber = "CF 5%";
            pane.CommitHeld();

            Assert.Equal(["MF 0.5%", "Only R21", "CF 5%"], PartsOf(session, "R21"));
            Assert.Equal(["MF 0.5%", "CF 5%"], PartsOf(session, "R22"));

            // Removed: from every one, the row only R21 has staying.
            Dispatcher.UIThread.RunJobs();
            Button remove = Get<ItemsControl>(pane, "ReplacementsItemsControl")
                .GetRealizedContainers()
                .SelectMany(container => container.GetVisualDescendants().OfType<Button>())
                .First(button => button.Content as string == "Remove");
            Click(remove);
            pane.CommitHeld();

            Assert.Equal(["Only R21", "CF 5%"], PartsOf(session, "R21"));
            Assert.Equal(["CF 5%"], PartsOf(session, "R22"));
            Assert.Equal([], PartsOf(session, "U8"));

            window.Close();
        });
    }

    [Fact]
    public void Make_all_like_gives_every_one_what_the_one_picked_has_on_the_page_shown_as_one_step()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.ShowPage(ComponentPage.Replacements);

            // The button names the one, a "*1" pointing at the note under it: only this page changes (owner request,
            // 2026-11-09).
            Button make = Get<Button>(pane, "MakeAllLikeButton");
            Assert.True(Get<StackPanel>(pane, "MakeAllLikePanel").IsVisible);
            Assert.Equal("Make all have the same data as [R21] *1", pane.MakeAllLikeButtonTextForTests);

            TextBlock note = Get<TextBlock>(pane, "MakeAllLikeNote");
            Assert.True(note.IsVisible);
            Assert.Equal(
                "*1 The button changes only the data on this page, \"Replacement parts\" - the components' other pages stay as they are.",
                BoldRuns.TextOf(note));

            Click(make);

            Assert.Equal(["MF 1%", "Only R21"], PartsOf(session, "R22"));
            Assert.Equal("\"Replacement parts\" made like [R21]", session.UndoDescription);
            Assert.Equal(2, ReplacementRows(pane).Count);

            // Alike now: nothing to say about what differs.
            Assert.False(Get<TextBlock>(pane, "SeveralPageText").IsVisible);

            Assert.True(session.Undo());
            Assert.Equal(["MF 1%"], PartsOf(session, "R22"));

            window.Close();
        });
    }

    // A baseline is measured on one component; pins only while they share a package.
    [Fact]
    public void Oscilloscope_baselines_are_one_components_and_pins_only_while_the_package_is_shared()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Baselines);
            Assert.False(Get<StackPanel>(pane, "BaselinesSection").IsVisible);
            Assert.False(Get<StackPanel>(pane, "MakeAllLikePanel").IsVisible);
            Assert.Contains("measured on one component", Get<TextBlock>(pane, "SeveralPageText").Text, StringComparison.Ordinal);

            pane.CommitHeld();
            Assert.Single(session.Content.Board.ComponentImages);

            pane.ShowPage(ComponentPage.Pinout);
            Assert.True(Get<StackPanel>(pane, "PinoutSection").IsVisible);

            pane.ShowSeveral(session, ["R21", "U8"], string.Empty, Published());
            Dispatcher.UIThread.RunJobs();
            Assert.False(Get<StackPanel>(pane, "MakeAllLikePanel").IsVisible);
            Assert.Contains("packages differ (Axial, DIP-40)", Get<TextBlock>(pane, "SeveralPageText").Text, StringComparison.Ordinal);

            // Their pins are held back - the package box stays, where one package is given to all of them (code review,
            // 2026-11-09: the whole page went, and the bar sent them to "Basics", which has no such box).
            Assert.True(Get<StackPanel>(pane, "PinoutSection").IsVisible);
            Assert.False(Get<Grid>(pane, "PinsArea").IsVisible);
            Assert.True(Get<AutoCompleteBox>(pane, "PackageBox").IsEffectivelyVisible);

            Get<AutoCompleteBox>(pane, "PackageBox").Text = "DIP-40";
            pane.CommitHeld();

            Assert.Equal("DIP-40", Row(session, "R21").Package);
            Assert.Equal("DIP-40", Row(session, "U8").Package);
            Assert.True(Get<Grid>(pane, "PinsArea").IsVisible);

            window.Close();
        });
    }

    // "Add the N pins of DIP-14" is never offered for several (code review, 2026-11-09): the pins listed are only those all
    // of them have, so U1's own pins 1-14 read as missing and were added to it a second time.
    [Fact]
    public void Adding_the_case_pins_is_not_offered_for_several()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "NAND", Package = "DIP-14" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC", FriendlyName = "NOR", Package = "DIP-14" },
            ],
            ComponentPins = Enumerable.Range(1, 14)
                .Select(pin => new ComponentPinEntry { BoardLabel = "U1", Pin = pin.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = $"P{pin}" })
                .ToList(),
        };

        CachedWorkbooks.Write(boardFile, board);
        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.ShowSeveral(session, ["U1", "U2"], string.Empty, board);
            Dispatcher.UIThread.RunJobs();
            pane.ShowPage(ComponentPage.Pinout);

            Assert.True(Get<StackPanel>(pane, "PinoutSection").IsVisible);
            Assert.False(Get<Button>(pane, "AddCasePinsButton").IsVisible);

            pane.AddCasePins();
            pane.CommitHeld();

            Assert.Equal(14, ComponentPaneEdits.PinsOf(session.Content.Board, "U1", string.Empty).Count);
            Assert.Empty(ComponentPaneEdits.PinsOf(session.Content.Board, "U2", string.Empty));
            Assert.False(session.CanUndo);

            // One of them alone is offered it as before.
            pane.Show(session, "U2", string.Empty, board);
            Dispatcher.UIThread.RunJobs();
            Assert.True(Get<Button>(pane, "AddCasePinsButton").IsVisible);

            window.Close();
        });
    }

    // The KiCad nets are one footprint's: never shown beside the pins several share (code review, 2026-11-09: the first
    // one's nets stood on every shared pin), and "Fill pins from KiCad" fills nothing for several.
    [Fact]
    public void One_components_KiCad_nets_are_not_shown_for_several()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "NAND", Package = "DIP-14" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC", FriendlyName = "NOR", Package = "DIP-14" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U1", Pin = "14", Name = "VCC" },
                new ComponentPinEntry { BoardLabel = "U2", Pin = "14", Name = "VCC" },
            ],
        };

        CachedWorkbooks.Write(boardFile, board);
        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.ShowSeveral(session, ["U1", "U2"], string.Empty, board);
            Dispatcher.UIThread.RunJobs();
            pane.ShowPage(ComponentPage.Pinout);

            pane.UseKiCadPins(session, "U1", [new KiCadPin("14", "VCC", "Power", "+5V"), new KiCadPin("1", "A", "Input", "/A1")]);
            Dispatcher.UIThread.RunJobs();

            Assert.Single(pane.PinsForTests);
            Assert.All(pane.PinsForTests, row => Assert.False(row.HasNet));
            Assert.False(Get<TextBlock>(pane, "PinNetHeading").IsVisible);
            Assert.False(Get<Button>(pane, "FillFromKiCadButton").IsVisible);

            Assert.False(pane.FillFromKiCad());
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    // Categories that differ are no category missing; a list with nothing in common says so, not that there is nothing.
    [Fact]
    public void What_differs_is_never_said_to_be_missing()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "R21", "U8");

            Assert.Equal("Differs - 2 values", Get<ComboBox>(pane, "CategoryBox").PlaceholderText);
            Assert.False(Get<TextBlock>(pane, "CategoryRequiredText").IsVisible);

            pane.ShowPage(ComponentPage.Replacements);
            Assert.Empty(ReplacementRows(pane));
            Assert.True(Get<TextBlock>(pane, "NoReplacementsText").IsVisible);
            Assert.Equal("None that all of them have.", Get<TextBlock>(pane, "NoReplacementsText").Text);

            pane.Show(session, "U8", string.Empty, Published());
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("No replacement parts yet.", Get<TextBlock>(pane, "NoReplacementsText").Text);

            window.Close();
        });
    }

    // The line of the list under a field they differ in that offers `value`.
    private static Button DifferingLine(ComponentPane pane, string value)
    {
        Dispatcher.UIThread.RunJobs();

        return Get<ItemsControl>(pane, "DiffersItems")
            .GetRealizedContainers()
            .SelectMany(container => container.GetVisualDescendants().OfType<Button>().Prepend(container as Button).OfType<Button>())
            .First(button => button.Tag is ComponentPaneDifferingValue picked && picked.Value == value);
    }

    // ###########################################################################################
    // A FIELD THEY DIFFER IN LISTS THEIR VALUES UNDER IT WHILE IT HAS THE FOCUS - NEVER UNDER THE POINTER - AND WHAT IS TYPED
    // NARROWS IT (owner requests, 2026-11-09: "I do not want that mouse hover will trigger that list of entries, as it makes it
    // hard to navigate in the other fields. Only when cursor is active in a field it should show the list ... where I can start
    // typing also and then it will filter the list"). Typed into, it waits for a pick, Enter or the box being left - no
    // typing pause sets half a word on all of them; the arrow keys and Enter pick as in a drop-down's list. A field they
    // share lists nothing; the focus gone, the list goes.
    // ###########################################################################################
    [Fact]
    public void A_field_they_differ_in_lists_their_values_while_it_has_the_focus_narrowed_to_what_is_typed()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            TextBox friendly = Get<TextBox>(pane, "FriendlyNameBox");
            TextBox technical = Get<TextBox>(pane, "TechnicalNameBox");

            // The pointer over it lists nothing.
            window.MouseMove(friendly.TranslatePoint(new Avalonia.Point(10, 10), window)!.Value);
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.AreDifferingValuesShownForTests);

            // A field they share, focused: nothing either.
            technical.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.AreDifferingValuesShownForTests);

            friendly.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.AreDifferingValuesShownForTests);
            Assert.Equal(["Pull-up", "Pull-down"], pane.DifferingValuesShownForTests.Select(value => value.Shown).ToArray());
            Assert.Equal(["R21", "R22"], pane.DifferingValuesShownForTests.Select(value => value.Whose).ToArray());

            // Typed: the list narrows, and nothing goes into the draft by a pause.
            window.KeyTextInput("down");
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["Pull-down"], pane.DifferingValuesShownForTests.Select(value => value.Value).ToArray());
            Assert.False(pane.IsTypingPauseRunningForTests);
            Assert.Equal("Pull-up", Row(session, "R21").FriendlyName);

            // Nothing left to show, the list goes - and comes back as the text finds something again.
            window.KeyTextInput("x");
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.AreDifferingValuesShownForTests);
            window.KeyPressQwerty(PhysicalKey.Backspace, RawInputModifiers.None);
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.AreDifferingValuesShownForTests);

            // Down lights the first line, Enter sets it on all of them - one step.
            window.KeyPressQwerty(PhysicalKey.ArrowDown, RawInputModifiers.None);
            Assert.Equal(0, pane.DifferingValueLitForTests);
            Assert.Contains("Lit", DifferingLine(pane, "Pull-down").Classes);
            window.KeyPressQwerty(PhysicalKey.Enter, RawInputModifiers.None);
            Dispatcher.UIThread.RunJobs();

            Assert.False(pane.AreDifferingValuesShownForTests);
            Assert.Equal("Pull-down", Row(session, "R21").FriendlyName);
            Assert.Equal("Pull-down", Row(session, "R22").FriendlyName);
            Assert.Equal("Friendly name of [R21, R22]", session.UndoDescription);
            Assert.Equal("Pull-down", friendly.Text);

            // Alike now, the field lists nothing as it is focused again.
            technical.Focus();
            friendly.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.AreDifferingValuesShownForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // ANOTHER BOX'S TYPING PAUSE TAKES NOTHING TYPED IN A FIELD THEY DIFFER IN (code review, 2026-11-09): a pause started by
    // a shared field - Avalonia posts TextChanged, so it can start after the focus moved on - ended while "Res" was being
    // typed to narrow the list, and set "Res" on all of them. What is typed there waits for Enter, a pick or the box left.
    // ###########################################################################################
    [Fact]
    public void Another_boxs_typing_pause_takes_nothing_typed_in_a_field_they_differ_in()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            TextBox friendly = Get<TextBox>(pane, "FriendlyNameBox");
            TextBox technical = Get<TextBox>(pane, "TechnicalNameBox");

            technical.Text = "4K7 1%";
            friendly.Focus();
            Dispatcher.UIThread.RunJobs();
            window.KeyTextInput("Res");
            Dispatcher.UIThread.RunJobs();

            // The pause ends: the shared field goes in, the differing one waits.
            pane.CommitHeld(fromTyping: true);

            Assert.Equal("4K7 1%", Row(session, "R21").TechnicalNameOrValue);
            Assert.Equal("Pull-up", Row(session, "R21").FriendlyName);
            Assert.Equal("Pull-down", Row(session, "R22").FriendlyName);

            // Left, it goes in - on all of them.
            technical.Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Res", Row(session, "R21").FriendlyName);
            Assert.Equal("Res", Row(session, "R22").FriendlyName);

            window.Close();
        });
    }

    // ###########################################################################################
    // A CATEGORY PICKED FROM THE VALUES LISTED IS SET AS PICKED, WHATEVER ITS CAPITALS (code review, 2026-11-09): the box lists
    // each category's first spelling, so "ic" picked selected nothing in it, read back as no category - and every component
    // picked lost its own.
    // ###########################################################################################
    [Fact]
    public void A_category_picked_from_the_values_listed_is_set_whatever_its_capitals()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "A" },
                new ComponentEntry { BoardLabel = "U2", Category = "ic", FriendlyName = "B" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "C" },
            ],
        };

        CachedWorkbooks.Write(boardFile, board);
        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U1", "U2", "C1");

            Get<ComboBox>(pane, "CategoryBox").Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.AreDifferingValuesShownForTests);

            // Their categories differ: none is missing, so "Basics" is no page that stops Submit (code review, 2026-11-09: it
            // went red, the box left empty on purpose).
            Assert.False(pane.PagesShown![ComponentPage.Basics].HasError);
            Assert.False(Get<TextBlock>(pane, "CategoryRequiredText").IsVisible);

            Click(DifferingLine(pane, "ic"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["ic", "ic", "ic"], new[] { "U1", "U2", "C1" }.Select(label => Row(session, label).Category));

            window.Close();
        });
    }

    // ###########################################################################################
    // THE VALUES LISTED ARE AS TALL AS THEY MAY BE, THEN SCROLLED (code review, 2026-11-09: forty values ran off the screen, the
    // ones below its edge out of reach) - and a line lit by the keys is brought into view.
    // ###########################################################################################
    [Fact]
    public void Many_values_listed_scroll_and_the_line_lit_comes_into_view()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        string[] labels = Enumerable.Range(1, 30).Select(number => $"R{number}").ToArray();
        BoardData board = new()
        {
            Components = [.. labels.Select(label => new ComponentEntry { BoardLabel = label, Category = "Resistor", FriendlyName = $"Pull-up of {label}" })],
        };

        CachedWorkbooks.Write(boardFile, board);
        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, labels);

            Get<TextBox>(pane, "FriendlyNameBox").Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.AreDifferingValuesShownForTests);

            ScrollViewer scroller = Get<ScrollViewer>(pane, "DiffersScroller");
            Assert.True(scroller.Bounds.Height <= 320, $"The list is {scroller.Bounds.Height}px tall");
            Assert.True(scroller.Extent.Height > scroller.Bounds.Height, "Thirty values fit without scrolling");

            for (int line = 0; line < 30; line++)
            {
                window.KeyPressQwerty(PhysicalKey.ArrowDown, RawInputModifiers.None);
            }

            Dispatcher.UIThread.RunJobs();
            Assert.Equal(29, pane.DifferingValueLitForTests);
            Assert.True(scroller.Offset.Y > 0, "The line lit was not brought into view");

            // Picked: one step, named after the first few and how many more (code review, 2026-11-09: named after all thirty).
            window.KeyPressQwerty(PhysicalKey.Enter, RawInputModifiers.None);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("Friendly name of [R1, R2, R3, R4, R5, R6 and 24 more]", session.UndoDescription);

            window.Close();
        });
    }

    // ###########################################################################################
    // IN "PURPOSE OR NOTE", A NOTE OF SEVERAL LINES, DOWN MOVES THE CARET UNTIL IT STANDS AT THE NOTE'S END (code review,
    // 2026-11-09: the list open under it took Down, and the caret could not leave the first line) - then it reaches the list.
    // ###########################################################################################
    [Fact]
    public void In_a_note_of_several_lines_Down_moves_the_caret_before_it_reaches_the_values_listed()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "R21", Category = "Resistor", Purpose = "a\nb" },
                new ComponentEntry { BoardLabel = "R22", Category = "Resistor", Purpose = "c" },
            ],
        };

        CachedWorkbooks.Write(boardFile, board);
        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "R21", "R22");
            RichNoteTextBox purpose = Get<RichNoteTextBox>(pane, "PurposeBox");

            // Two lines typed, the second still empty - "a" narrows the list to R21's.
            purpose.Focus();
            purpose.Markup = "a\n";
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.AreDifferingValuesShownForTests);

            purpose.CaretIndex = 0;
            window.KeyPressQwerty(PhysicalKey.ArrowDown, RawInputModifiers.None);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(-1, pane.DifferingValueLitForTests);

            purpose.CaretIndex = purpose.Text!.Length;
            window.KeyPressQwerty(PhysicalKey.ArrowDown, RawInputModifiers.None);
            Assert.Equal(0, pane.DifferingValueLitForTests);

            window.Close();
        });
    }

    // A line clicked is set on all of them, the box keeping the focus (the lines take none); leaving the box closes the list.
    [Fact]
    public void A_value_clicked_in_the_list_is_set_on_all_of_them_and_leaving_the_box_closes_the_list()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            TextBox friendly = Get<TextBox>(pane, "FriendlyNameBox");
            TextBox technical = Get<TextBox>(pane, "TechnicalNameBox");

            friendly.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.AreDifferingValuesShownForTests);
            Assert.False(DifferingLine(pane, "Pull-up").Focusable);

            // Left: the list goes.
            technical.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.AreDifferingValuesShownForTests);
            Assert.Equal("Pull-up", Row(session, "R21").FriendlyName);

            friendly.Focus();
            Dispatcher.UIThread.RunJobs();
            Click(DifferingLine(pane, "Pull-down"));

            Assert.False(pane.AreDifferingValuesShownForTests);
            Assert.Equal("Pull-down", Row(session, "R21").FriendlyName);
            Assert.Equal("Pull-down", Row(session, "R22").FriendlyName);
            Assert.Equal("Friendly name of [R21, R22]", session.UndoDescription);

            window.Close();
        });
    }

    // ###########################################################################################
    // "THIS DATA IS DIFFERENT PER REGION" STAYS WITH SEVERAL PICKED (owner request, 2026-11-09: "Can't it still have this,
    // even if selecting multiple components?"): it says all of them, none or some (half-ticked), and a click makes all of
    // them so - one step.
    // ###########################################################################################
    [Fact]
    public void The_per_region_tick_says_all_none_or_some_and_a_click_makes_all_of_them_so()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        BoardData published = Published().WithRegions(["PAL", "NTSC"]);
        CachedWorkbooks.Write(boardFile, published);
        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.ShowSeveral(session, Both, "PAL", published);
            Dispatcher.UIThread.RunJobs();

            CheckBox tick = Get<CheckBox>(pane, "BasicsRegionalCheck");
            Assert.True(tick.IsVisible);
            Assert.False(tick.IsChecked);

            tick.IsChecked = true;
            tick.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.True(ComponentPaneEdits.IsRegional(session.Content.Board, "R21"));
            Assert.True(ComponentPaneEdits.IsRegional(session.Content.Board, "R22"));
            Assert.False(ComponentPaneEdits.IsRegional(session.Content.Board, "U8"));
            Assert.Contains("R21, R22", session.UndoDescription, StringComparison.Ordinal);

            pane.ShowSeveral(session, Both, "PAL", published);
            Dispatcher.UIThread.RunJobs();
            Assert.True(tick.IsChecked);

            // One of them told apart and one not: half-ticked, and a click ticks them all.
            Assert.True(session.Undo());
            session.Apply("R21 per region", content => content with { Board = ComponentPaneEdits.WithRegionalRows(content.Board, "R21", "PAL", true) });
            pane.ShowSeveral(session, Both, "PAL", published);
            Dispatcher.UIThread.RunJobs();
            Assert.Null(tick.IsChecked);

            tick.IsChecked = false;
            tick.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.True(ComponentPaneEdits.IsRegional(session.Content.Board, "R22"));

            window.Close();
        });
    }

    // One picked again: the pane is that one's, as ever.
    [Fact]
    public void One_shown_again_is_the_pane_of_one_component()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.Show(session, "R21", string.Empty, Published());
            Dispatcher.UIThread.RunJobs();

            Assert.False(pane.IsSeveral);
            Assert.False(Get<Border>(pane, "SeveralBar").IsVisible);
            Assert.True(Get<TextBox>(pane, "BoardLabelBox").IsEnabled);
            Assert.Equal("R21", Get<TextBox>(pane, "BoardLabelBox").Text);
            Assert.Equal("Pull-up", Get<TextBox>(pane, "FriendlyNameBox").Text);
            Assert.NotEqual("Differs - 2 values", Get<TextBox>(pane, "FriendlyNameBox").PlaceholderText);
            Assert.True(Get<WrapPanel>(pane, "ActionsPanel").IsVisible);
            Assert.True(Get<Button>(pane, "DeleteComponentButton").IsVisible);

            window.Close();
        });
    }
}
