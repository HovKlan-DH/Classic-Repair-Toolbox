using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The category chips and state pills in the full worklog editor window - the one place in the app
// that offers them, since the quick "Create worklog" card that used to mirror them is gone and
// drawing an area now opens this editor directly.
//
// The state pill used to be outline-only when selected, which on the pale panel background left
// "selected" and "unselected" separated by little more than a 1px border-width difference - the
// selected pill was genuinely hard to pick out. It is now filled, like the category chips.
//
// The category icons are asserted by CODEPOINT, not just by presence. Font Awesome ships the Free
// Regular face as a 362-glyph subset, so a glyph that exists in Solid is often absent from
// Regular; picking a codepoint from memory produces a tofu box that no test would otherwise
// catch. Each value below was read out of the shipped OTF.
[Collection("HeadlessUi")]
public class WorklogCategoryAndStateVisualsTests
{
    // Verified against Assets/Fonts with fontTools: file-lines U+F15C is present in the Free
    // Regular subset (internal glyph name "i59"), spray-can-sparkles and bug are Solid-only.
    private const string NoteGlyph = "\uF15C";
    private const string CosmeticGlyph = "\uF5D0";
    private const string IssueGlyph = "\uF188";

    // fa-solid lock-open / lock, read out of the shipped Solid OTF.
    private const string OpenGlyph = "\uF3C1";

    private const string ClosedGlyph = "\uF023";

    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static WorklogEntryRecord CreateEntry(string state) => new()
    {
        Id = 7,
        SchematicName = "Sch",
        Title = "Bad cap",
        Category = "Note",
        State = state,
        AreaX = 10,
        AreaY = 10,
        AreaWidth = 50,
        AreaHeight = 50,
    };

    private static void WithEditor(string state, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, CreateEntry(state), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
            }
        });
    }

    private static Color? SolidColorOf(IBrush? brush) => (brush as ISolidColorBrush)?.Color;

    // ------------------------------------------------------------- state pill selection cue

    // The selected pill is FILLED with its state colour: its background must be the same colour
    // as its border, which is what distinguishes a filled pill from an outlined one regardless of
    // which theme is active.
    [Theory]
    [InlineData("Open", "EditorStateOpenPill", "EditorStateClosedPill")]
    [InlineData("Closed", "EditorStateClosedPill", "EditorStateOpenPill")]
    public void The_selected_state_pill_is_filled_and_the_other_is_not(string state, string selectedName, string unselectedName)
    {
        WithEditor(state, window =>
        {
            var selected = window.FindControl<Border>(selectedName)!;
            var unselected = window.FindControl<Border>(unselectedName)!;

            var selectedBg = SolidColorOf(selected.Background);
            var selectedBorder = SolidColorOf(selected.BorderBrush);

            Assert.NotNull(selectedBg);
            Assert.NotNull(selectedBorder);
            Assert.Equal(selectedBorder!.Value, selectedBg!.Value);

            // The unselected pill keeps the neutral form background, which must NOT match its own
            // border - otherwise both pills would read as filled and nothing would mark the choice.
            var unselectedBg = SolidColorOf(unselected.Background);
            var unselectedBorder = SolidColorOf(unselected.BorderBrush);

            Assert.NotNull(unselectedBg);
            Assert.NotNull(unselectedBorder);
            Assert.NotEqual(unselectedBorder!.Value, unselectedBg!.Value);
        });
    }

    // The padlock goes white inside the filled pill. One left at its state colour would sit on a
    // fill of that same colour and simply disappear.
    [Theory]
    [InlineData("Open", "EditorStateOpenDot", "EditorStateClosedDot")]
    [InlineData("Closed", "EditorStateClosedDot", "EditorStateOpenDot")]
    public void The_selected_pills_icon_turns_white_while_the_other_keeps_its_state_colour(string state, string selectedIconName, string unselectedIconName)
    {
        WithEditor(state, window =>
        {
            var selectedIcon = window.FindControl<TextBlock>(selectedIconName)!;
            var unselectedIcon = window.FindControl<TextBlock>(unselectedIconName)!;

            Assert.Equal(Colors.White, SolidColorOf(selectedIcon.Foreground));
            Assert.NotEqual(Colors.White, SolidColorOf(unselectedIcon.Foreground));
        });
    }

    // The padlocks themselves: open = fa-solid lock-open, closed = fa-solid lock. Asserted by
    // codepoint, because a wrong one renders a blank box that no other test would catch.
    [Fact]
    public void The_state_pills_use_the_open_and_closed_padlocks()
    {
        WithEditor("Open", window =>
        {
            Assert.Equal(OpenGlyph, window.FindControl<TextBlock>("EditorStateOpenDot")!.Text);
            Assert.Equal(ClosedGlyph, window.FindControl<TextBlock>("EditorStateClosedDot")!.Text);
        });
    }

    // The selected pill must be visibly distinct from the unselected one. Asserting the
    // backgrounds differ is the property that actually failed before: outline-only left both
    // pills on near-identical backgrounds.
    [Theory]
    [InlineData("Open")]
    [InlineData("Closed")]
    public void The_two_state_pills_never_share_a_background(string state)
    {
        WithEditor(state, window =>
        {
            var openBg = SolidColorOf(window.FindControl<Border>("EditorStateOpenPill")!.Background);
            var closedBg = SolidColorOf(window.FindControl<Border>("EditorStateClosedPill")!.Background);

            Assert.NotEqual(openBg, closedBg);
        });
    }

    // ------------------------------------------------------------- category icons

    [Fact]
    public void The_editor_category_chips_carry_their_font_awesome_icons()
    {
        WithEditor("Open", window =>
        {
            Assert.Equal(NoteGlyph, window.FindControl<TextBlock>("EditorCategoryNoteIcon")!.Text);
            Assert.Equal(CosmeticGlyph, window.FindControl<TextBlock>("EditorCategoryCosmeticIcon")!.Text);
            Assert.Equal(IssueGlyph, window.FindControl<TextBlock>("EditorCategoryIssueIcon")!.Text);
        });
    }

    // Note is the one icon taken from the Regular face; the other two only exist in Solid. Getting
    // this pairing wrong renders a blank box rather than failing, so it is pinned explicitly.
    [Fact]
    public void The_note_icon_uses_the_regular_face_and_the_others_use_solid()
    {
        WithEditor("Open", window =>
        {
            string note = window.FindControl<TextBlock>("EditorCategoryNoteIcon")!.FontFamily.Name;
            string cosmetic = window.FindControl<TextBlock>("EditorCategoryCosmeticIcon")!.FontFamily.Name;
            string issue = window.FindControl<TextBlock>("EditorCategoryIssueIcon")!.FontFamily.Name;

            Assert.NotEqual(note, cosmetic);
            Assert.Equal(cosmetic, issue);
        });
    }

    // ------------------------------------------------------------- icon colour

    // The icon takes its LABEL's colour, not a colour of its own. Left at one fixed foreground it
    // would either disappear into the selected chip's filled background or stay dark while its own
    // label went white - a half-coloured chip either way.
    [Theory]
    [InlineData("Note", "EditorCategoryNoteIcon", "EditorCategoryNoteText")]
    [InlineData("Cosmetic", "EditorCategoryCosmeticIcon", "EditorCategoryCosmeticText")]
    [InlineData("Issue", "EditorCategoryIssueIcon", "EditorCategoryIssueText")]
    public void A_category_icon_always_matches_its_own_label_colour(string category, string iconName, string labelName)
    {
        WithEditorCategory(category, window =>
        {
            // Every chip is checked, selected and unselected alike - the rule is that icon and
            // label agree, which must hold in both states rather than only on the selected one.
            foreach (var (icon, label) in new[]
                     {
                         (iconName, labelName),
                         ("EditorCategoryNoteIcon", "EditorCategoryNoteText"),
                         ("EditorCategoryCosmeticIcon", "EditorCategoryCosmeticText"),
                         ("EditorCategoryIssueIcon", "EditorCategoryIssueText"),
                     })
            {
                var iconColor = SolidColorOf(window.FindControl<TextBlock>(icon)!.Foreground);
                var labelColor = SolidColorOf(window.FindControl<TextBlock>(label)!.Foreground);

                Assert.Equal(labelColor, iconColor);
            }
        });
    }

    // The selected chip is filled, so its icon must be white - the specific value the rule above
    // produces, asserted directly so "both are the same wrong colour" cannot pass.
    [Fact]
    public void The_selected_category_icon_is_white()
    {
        WithEditorCategory("Issue", window =>
        {
            Assert.Equal(Colors.White, SolidColorOf(window.FindControl<TextBlock>("EditorCategoryIssueIcon")!.Foreground));
            Assert.NotEqual(Colors.White, SolidColorOf(window.FindControl<TextBlock>("EditorCategoryNoteIcon")!.Foreground));
        });
    }

    // ------------------------------------------------------------- button wording

    [Fact]
    public void The_editor_commits_with_an_update_worklog_button()
    {
        WithEditorCategory("Note", window =>
            Assert.Equal("Update worklog", window.FindControl<Button>("EditorSaveButton")!.Content));
    }

    private static void WithEditorCategory(string category, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            var entry = CreateEntry("Open");
            entry.Category = category;

            using var bitmap = CreateBitmap();
            window.Initialize(1, entry, bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
            }
        });
    }
}
