using System.Globalization;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The seven collapsible list sections in the full editor - Links, Work done, Comments, Components
// in scope, Components completed, Photos and Files - plus the header layout that keeps their
// titles, counts and action links inside the panel at any width.
//
// The layout bug these guard: the headers were fixed-width rows, so a long title plus a count
// ("Components in scope   132 found   All") ran past the right-hand panel edge and the "All"/"None"
// links were clipped off-screen entirely - unreachable, not merely ugly.
[Collection("HeadlessUi")]
public class WorklogListSectionTests
{
    // fa-regular square-plus / square-minus, read out of the shipped OTF. Both exist in the Free
    // Regular face (as internal glyphs i43/i51) even though it is only a 362-glyph subset.
    private const string ExpandGlyph = "\uF0FE";
    private const string CollapseGlyph = "\uF146";

    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static WorklogEntryRecord CreateEntry() => new()
    {
        Id = 1,
        SchematicName = "Sch",
        Title = "Recap the board",
        Category = "Issue",
        State = "Open",
        AreaX = 10,
        AreaY = 10,
        AreaWidth = 50,
        AreaHeight = 50,
        ComponentLabels = new List<string> { "C1", "C2" },
    };

    // A deliberately large scope: 132 components, the count from the reported screenshot, so the
    // header carries the longest text it realistically has to fit.
    private static List<(string BoardLabel, string DisplayName)> LargeScope()
    {
        var scope = new List<(string, string)>();
        for (int i = 1; i <= 132; i++)
        {
            scope.Add(($"C{i}", "Ceramic | 100pF 25V"));
        }
        return scope;
    }

    private static void WithEditor(double width, Action<WorklogEntryEditorWindow> body, WorklogEntryRecord? entry = null)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = width;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, entry ?? CreateEntry(), bitmap);
            window.InitializeComponentScope(LargeScope());

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

    private static void ClickHeader(Window window, string headerTag)
    {
        var header = window.GetVisualDescendants()
            .OfType<Border>()
            .First(b => (b.Tag as string) == headerTag);

        // Scrolled into view first. The right-hand column is taller than its viewport, so the last
        // section (Files) sits below the visible area - a press at its coordinates lands on
        // whatever is clipping it and silently does nothing, which looks exactly like a broken
        // toggle. Verified: its header measured y=646 against a 626px viewport.
        header.BringIntoView();
        Dispatcher.UIThread.RunJobs();
        window.Measure(window.ClientSize);
        window.Arrange(new Rect(window.ClientSize));
        Dispatcher.UIThread.RunJobs();

        var centre = header.TranslatePoint(new Point(8, header.Bounds.Height / 2), window);
        Assert.NotNull(centre);

        window.MouseDown(centre!.Value, MouseButton.Left);
        window.MouseUp(centre.Value, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static readonly (string HeaderTag, string IconName, string BodyName)[] Sections =
    {
        ("EditorLinksHeader", "EditorLinksHeaderIcon", "EditorLinksList"),
        ("EditorWorkDoneHeader", "EditorWorkDoneHeaderIcon", "EditorWorkDoneList"),
        ("EditorCommentsHeader", "EditorCommentsHeaderIcon", "EditorCommentsList"),
        ("EditorComponentScopeHeader", "EditorComponentScopeHeaderIcon", "EditorComponentScopeBody"),
        ("EditorComponentCompletedHeader", "EditorComponentCompletedHeaderIcon", "EditorComponentCompletedBody"),
        ("EditorPhotosHeader", "EditorPhotosHeaderIcon", "EditorPhotosList"),
        ("EditorFilesHeader", "EditorFilesHeaderIcon", "EditorFilesList"),
    };

    public static IEnumerable<object[]> AllSections() =>
        Sections.Select(s => new object[] { s.HeaderTag, s.IconName, s.BodyName });

    // ------------------------------------------------------------- collapse / expand

    // Every list starts open, showing the collapse ("minus") icon - a worklog that opened with
    // everything folded away would hide content the user never chose to hide.
    [Theory]
    [MemberData(nameof(AllSections))]
    public void Every_section_starts_expanded_showing_the_collapse_icon(string headerTag, string iconName, string bodyName)
    {
        _ = headerTag;

        WithEditor(1200, window =>
        {
            Assert.Equal(CollapseGlyph, window.FindControl<TextBlock>(iconName)!.Text);
            Assert.True(window.FindControl<Control>(bodyName)!.IsVisible);
        });
    }

    [Theory]
    [MemberData(nameof(AllSections))]
    public void Clicking_a_header_collapses_its_list_and_swaps_the_icon(string headerTag, string iconName, string bodyName)
    {
        WithEditor(1200, window =>
        {
            ClickHeader(window, headerTag);

            Assert.Equal(ExpandGlyph, window.FindControl<TextBlock>(iconName)!.Text);
            Assert.False(window.FindControl<Control>(bodyName)!.IsVisible);
        });
    }

    [Theory]
    [MemberData(nameof(AllSections))]
    public void Clicking_a_collapsed_header_expands_it_again(string headerTag, string iconName, string bodyName)
    {
        WithEditor(1200, window =>
        {
            ClickHeader(window, headerTag);
            ClickHeader(window, headerTag);

            Assert.Equal(CollapseGlyph, window.FindControl<TextBlock>(iconName)!.Text);
            Assert.True(window.FindControl<Control>(bodyName)!.IsVisible);
        });
    }

    // Sections are independent - folding one must not fold its neighbours.
    [Fact]
    public void Collapsing_one_section_leaves_the_others_open()
    {
        WithEditor(1200, window =>
        {
            ClickHeader(window, "EditorCommentsHeader");

            Assert.False(window.FindControl<Control>("EditorCommentsList")!.IsVisible);
            Assert.True(window.FindControl<Control>("EditorLinksList")!.IsVisible);
            Assert.True(window.FindControl<Control>("EditorPhotosList")!.IsVisible);
        });
    }

    // A collapsed section hides its "No links added" line too - otherwise folding an empty list
    // leaves its empty-state text floating under a closed header.
    [Fact]
    public void Collapsing_an_empty_section_hides_its_empty_state_line()
    {
        WithEditor(1200, window =>
        {
            Assert.True(window.FindControl<TextBlock>("EditorNoLinksText")!.IsVisible);

            ClickHeader(window, "EditorLinksHeader");

            Assert.False(window.FindControl<TextBlock>("EditorNoLinksText")!.IsVisible);
        });
    }

    // ...and expanding brings it back, because the list really is empty.
    [Fact]
    public void Expanding_an_empty_section_restores_its_empty_state_line()
    {
        WithEditor(1200, window =>
        {
            ClickHeader(window, "EditorLinksHeader");
            ClickHeader(window, "EditorLinksHeader");

            Assert.True(window.FindControl<TextBlock>("EditorNoLinksText")!.IsVisible);
        });
    }

    // The trap in the other direction: expanding a NON-empty section must not reveal its
    // "No comments added" line, which would sit above a list that has comments in it. The empty
    // line is restored from the row count, not from the fold state.
    [Fact]
    public void Expanding_a_populated_section_does_not_show_its_empty_state_line()
    {
        var entry = CreateEntry();
        entry.Comments.Add(new WorklogCommentRecord { Id = 1, Text = "Checked it", Date = DateTime.Now });

        WithEditor(1200, window =>
        {
            Assert.False(window.FindControl<TextBlock>("EditorNoCommentsText")!.IsVisible);

            ClickHeader(window, "EditorCommentsHeader");
            ClickHeader(window, "EditorCommentsHeader");

            Assert.False(window.FindControl<TextBlock>("EditorNoCommentsText")!.IsVisible);
        }, entry);
    }

    // ------------------------------------------------------------- totals

    [Fact]
    public void An_empty_list_reports_none_rather_than_a_zero()
    {
        WithEditor(1200, window =>
        {
            Assert.Equal("none", window.FindControl<TextBlock>("EditorLinksCountText")!.Text);
            Assert.Equal("none", window.FindControl<TextBlock>("EditorPhotosCountText")!.Text);
            Assert.Equal("none", window.FindControl<TextBlock>("EditorFilesCountText")!.Text);
        });
    }

    // Singular and plural, so a one-item list does not read "1 comments".
    [Fact]
    public void A_single_item_is_counted_in_the_singular()
    {
        var entry = CreateEntry();
        entry.Comments.Add(new WorklogCommentRecord { Id = 1, Text = "One", Date = DateTime.Now });

        WithEditor(1200, window =>
            Assert.Equal("1 comment", window.FindControl<TextBlock>("EditorCommentsCountText")!.Text), entry);
    }

    [Fact]
    public void Several_items_are_counted_in_the_plural()
    {
        var entry = CreateEntry();
        entry.Comments.Add(new WorklogCommentRecord { Id = 1, Text = "One", Date = DateTime.Now });
        entry.Comments.Add(new WorklogCommentRecord { Id = 2, Text = "Two", Date = DateTime.Now });

        WithEditor(1200, window =>
            Assert.Equal("2 comments", window.FindControl<TextBlock>("EditorCommentsCountText")!.Text), entry);
    }

    // Work done keeps its hours and cost totals, which were previously built into the title itself.
    // They moved to the count so the title stays a plain title, like every other section.
    [Fact]
    public void Work_done_reports_its_hours_and_cost_alongside_the_count()
    {
        var entry = CreateEntry();
        entry.WorkDoneItems.Add(new WorklogWorkDoneRecord { Id = 1, Text = "Recap", Date = DateTime.Now, HoursSpent = 2, Cost = 30 });
        entry.WorkDoneItems.Add(new WorklogWorkDoneRecord { Id = 2, Text = "Test", Date = DateTime.Now, HoursSpent = 1.5, Cost = 0 });

        WithEditor(1200, window =>
        {
            string text = window.FindControl<TextBlock>("EditorWorkDoneCountText")!.Text!;

            Assert.Contains("2 entries", text);

            // The time as HOURS AND MINUTES, not the decimal hours it is stored as - see
            // WorklogDurationFormatter. That also settles what used to be a locale problem here:
            // "3.5" needed the CURRENT culture's decimal separator to be asserted safely, while
            // "3 hours and 30 minutes" is two whole integers and reads the same everywhere.
            Assert.Contains("3 hours and 30 minutes", text);
            Assert.Contains("30", text);
        }, entry);
    }

    // The scope count reports how many are SELECTED out of how many are offered - "132 found" said
    // nothing about the choice the user had actually made.
    [Fact]
    public void The_scope_count_reports_the_selection_against_the_total()
    {
        WithEditor(1200, window =>
            Assert.Equal("2 of 132 selected", window.FindControl<TextBlock>("EditorComponentCountText")!.Text));
    }

    // ------------------------------------------------------------- header layout

    // THE reported bug, at the narrowest window the editor allows (MinWidth 900). Nothing in a
    // header may extend past the window's right edge - the counts used to be clipped and the
    // "All"/"None" links pushed off-screen entirely, where they could not be clicked at all.
    [Fact]
    public void Nothing_in_a_header_is_clipped_at_the_minimum_window_width()
    {
        WithEditor(900, window =>
        {
            double edge = window.ClientSize.Width;

            foreach (string name in new[] { "EditorComponentCountText", "EditorCompletedCountText", "EditorLinksCountText" })
            {
                var text = window.FindControl<TextBlock>(name)!;
                double right = text.TranslatePoint(new Point(text.Bounds.Width, 0), window)!.Value.X;

                Assert.True(right <= edge + 0.5, $"{name} ('{text.Text}') ends at {right}, past the window edge {edge}");
            }
        });
    }

    [Fact]
    public void The_all_and_none_links_stay_inside_the_window_at_the_minimum_width()
    {
        WithEditor(900, window =>
        {
            double edge = window.ClientSize.Width;

            foreach (string panelName in new[] { "EditorComponentScopePanel", "EditorComponentCompletedPanel" })
            {
                foreach (var button in window.FindControl<StackPanel>(panelName)!.GetVisualDescendants().OfType<Button>())
                {
                    if (button.Content is not string content)
                        continue;

                    double right = button.TranslatePoint(new Point(button.Bounds.Width, 0), window)!.Value.X;

                    Assert.True(right <= edge + 0.5, $"{panelName} '{content}' ends at {right}, past the window edge {edge}");
                }
            }
        });
    }

    // The links stay on the header's FIRST line even when the title and count wrap below them, so
    // they keep a fixed, predictable position rather than drifting down with the text.
    [Fact]
    public void The_all_and_none_links_stay_on_the_headers_first_line_when_the_title_wraps()
    {
        WithEditor(900, window =>
        {
            var title = window.GetVisualDescendants()
                .OfType<TextBlock>()
                .First(t => t.Text == "Components in scope");

            var count = window.FindControl<TextBlock>("EditorComponentCountText")!;
            var allButton = window.FindControl<StackPanel>("EditorComponentScopePanel")!
                .GetVisualDescendants()
                .OfType<Button>()
                .First(b => (b.Content as string) == "All");

            double titleY = title.TranslatePoint(new Point(0, 0), window)!.Value.Y;
            double countY = count.TranslatePoint(new Point(0, 0), window)!.Value.Y;
            double allY = allButton.TranslatePoint(new Point(0, 0), window)!.Value.Y;

            // The count really has wrapped below the title at this width - otherwise the assertion
            // underneath proves nothing.
            Assert.True(countY > titleY, $"the count did not wrap (title {titleY}, count {countY}) - this test is not exercising the case it claims");

            // ...and the link stayed up on the header's FIRST line rather than following the count
            // down to the second. It sits a few pixels above the title's top because the two have
            // different heights (16px title, 12px link) in a top-aligned row, so the test asserts
            // which LINE it is on rather than pixel equality with the title.
            Assert.True(allY < countY, $"the All link at {allY} dropped to the wrapped count line at {countY}");
            Assert.True(allY <= titleY + 6.0, $"the All link at {allY} is not on the title line at {titleY}");
        });
    }

    // ------------------------------------------------------------- persisted folds

    // The fold state is restored from the entry, so a section the user collapsed last time comes
    // back collapsed.
    [Fact]
    public void A_section_saved_as_collapsed_opens_collapsed()
    {
        var entry = CreateEntry();
        entry.CollapsedSections = new List<string> { "EditorCommentsHeader", "EditorPhotosHeader" };

        WithEditor(1200, window =>
        {
            Assert.False(window.FindControl<Control>("EditorCommentsList")!.IsVisible);
            Assert.False(window.FindControl<Control>("EditorPhotosList")!.IsVisible);

            // Everything else is unaffected - only the named sections fold.
            Assert.True(window.FindControl<Control>("EditorLinksList")!.IsVisible);
        }, entry);
    }

    // An entry written before the field existed - or one never folded - opens with everything
    // showing, because an absent key means "expanded".
    [Fact]
    public void An_entry_with_no_saved_folds_opens_fully_expanded()
    {
        var entry = CreateEntry();
        entry.CollapsedSections = new List<string>();

        WithEditor(1200, window =>
        {
            foreach (var (_, _, bodyName) in Sections)
            {
                Assert.True(window.FindControl<Control>(bodyName)!.IsVisible, $"{bodyName} did not open expanded");
            }
        }, entry);
    }

    // Restoring saved folds must not look like an edit - it is a reading preference, not a change
    // to the worklog, so "Update worklog" stays disabled.
    [Fact]
    public void Restoring_saved_folds_does_not_mark_the_window_dirty()
    {
        var entry = CreateEntry();
        entry.CollapsedSections = new List<string> { "EditorCommentsHeader" };

        WithEditor(1200, window =>
            Assert.False(window.FindControl<Button>("EditorSaveButton")!.IsEnabled), entry);
    }

    // ...and neither does folding one by hand. It saves itself immediately instead, the way the
    // sub-lists do, so closing with Cancel cannot lose it.
    [Fact]
    public void Collapsing_a_section_does_not_mark_the_window_dirty()
    {
        WithEditor(1200, window =>
        {
            ClickHeader(window, "EditorCommentsHeader");

            Assert.False(window.FindControl<Button>("EditorSaveButton")!.IsEnabled);
        });
    }

    // ------------------------------------------------------------- auto-expand on add

    // Changing a status writes an automatic "Worklog closed" comment, but must NOT unfold the
    // Comments section: the user was changing a status, not asking to read comments. Unfolding a
    // list they deliberately collapsed, every time they touch a pill, is the app second-guessing
    // them. Only a direct "Add comment" click expands - see below.
    [Fact]
    public void An_automatic_comment_leaves_a_collapsed_comments_section_collapsed()
    {
        var entry = CreateEntry();
        entry.CollapsedSections = new List<string> { "EditorCommentsHeader" };

        WithEditor(1200, window =>
        {
            Assert.False(window.FindControl<Control>("EditorCommentsList")!.IsVisible);

            ClickStatePill(window, "EditorStateClosedPill");

            Assert.False(window.FindControl<Control>("EditorCommentsList")!.IsVisible);
            Assert.Equal(ExpandGlyph, window.FindControl<TextBlock>("EditorCommentsHeaderIcon")!.Text);
        }, entry);
    }

    // ...and the comment really was written, so the test above is proving that a fold survived a
    // genuine addition rather than that nothing happened at all.
    [Fact]
    public void An_automatic_comment_is_still_recorded_while_the_section_stays_collapsed()
    {
        var entry = CreateEntry();
        entry.CollapsedSections = new List<string> { "EditorCommentsHeader" };

        WithEditor(1200, window =>
        {
            ClickStatePill(window, "EditorStateClosedPill");

            var rows = window.FindControl<ItemsControl>("EditorCommentsList")!.ItemsSource!
                .Cast<WorklogCommentRow>()
                .Select(r => r.Text)
                .ToList();

            Assert.Contains("Worklog closed", rows);
        }, entry);
    }

    // A section the user has open is left open by an automatic comment too - the rule is "do not
    // touch the fold", not "close it".
    [Fact]
    public void An_automatic_comment_leaves_an_open_comments_section_open()
    {
        WithEditor(1200, window =>
        {
            Assert.True(window.FindControl<Control>("EditorCommentsList")!.IsVisible);

            ClickStatePill(window, "EditorStateClosedPill");

            Assert.True(window.FindControl<Control>("EditorCommentsList")!.IsVisible);
        });
    }

    private static void ClickStatePill(Window window, string pillName)
    {
        var pill = window.FindControl<Border>(pillName)!;

        var centre = pill.TranslatePoint(new Point(pill.Bounds.Width / 2, pill.Bounds.Height / 2), window);
        Assert.NotNull(centre);

        window.MouseDown(centre!.Value, MouseButton.Left);
        window.MouseUp(centre.Value, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }
}
