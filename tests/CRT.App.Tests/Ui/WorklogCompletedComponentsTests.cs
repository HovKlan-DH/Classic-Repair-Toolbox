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

// "Mark components completed" - the progress checklist under "Mark components in scope", for
// tracking a job like "replace every capacitor on this board".
//
// Its rows mirror the TICKED rows of the scope list, so the two can never disagree about which
// components the entry covers. The rules that follow from that, and which these pin down:
//
//   - a component newly ticked INTO scope appears here UNTICKED. It is work still to do, which is
//     the point of the list; arriving pre-ticked would overstate progress.
//   - a component unticked OUT of scope disappears and loses its completed state. It is no longer
//     part of the entry, so a remembered "done" flag would be about work outside it.
//   - progress already made survives an unrelated scope edit - the rows are rebuilt, the ticks are
//     carried across.
[Collection("HeadlessUi")]
public class WorklogCompletedComponentsTests
{
    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static readonly (string BoardLabel, string DisplayName)[] Scope =
    {
        ("C1", "Capacitor | 100nF"),
        ("C2", "Capacitor | 220nF"),
        ("C3", "Capacitor | 470nF"),
    };

    private static WorklogEntryRecord CreateEntry(string[] inScope, string[]? completed = null) => new()
    {
        Id = 7,
        SchematicName = "Sch",
        Title = "Recap the board",
        Category = "Issue",
        State = "Open",
        AreaX = 10,
        AreaY = 10,
        AreaWidth = 50,
        AreaHeight = 50,
        ComponentLabels = inScope.ToList(),
        CompletedComponentLabels = (completed ?? Array.Empty<string>()).ToList(),
    };

    private static void WithEditor(WorklogEntryRecord entry, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            // Wide enough that the right-hand column's "All"/"None" links land INSIDE the window.
            // At 1000px they sat at x=1059 - clipped off-screen, so a mouse press at their centre
            // hit nothing and every link-driven test failed while the logic was perfectly fine.
            var window = new WorklogEntryEditorWindow();
            window.Width = 1400;
            window.Height = 800;

            using var bitmap = CreateBitmap();
            window.Initialize(1, entry, bitmap);
            window.InitializeComponentScope(Scope);

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

    private static List<WorklogEntryComponentRow> Rows(Window window, string listName) =>
        window.FindControl<ItemsControl>(listName)!.ItemsSource!.Cast<WorklogEntryComponentRow>().ToList();

    private static List<WorklogEntryComponentRow> ScopeRows(Window window) => Rows(window, "EditorComponentList");

    private static List<WorklogEntryComponentRow> CompletedRows(Window window) => Rows(window, "EditorCompletedComponentList");

    private static List<string> CompletedLabels(Window window) =>
        CompletedRows(window).Select(r => r.BoardLabel).ToList();

    // Clicks a scope row the way a user does, so the real handler (and the rebuild it triggers)
    // runs rather than the row's IsChecked being poked directly.
    private static void ClickScopeRow(Window window, string boardLabel)
    {
        var list = window.FindControl<ItemsControl>("EditorComponentList")!;

        var row = list.GetVisualDescendants()
            .OfType<Border>()
            .First(b => b.DataContext is WorklogEntryComponentRow r && r.BoardLabel == boardLabel);

        var centre = row.TranslatePoint(new Point(row.Bounds.Width / 2, row.Bounds.Height / 2), window);
        Assert.NotNull(centre);

        window.MouseDown(centre!.Value, MouseButton.Left);
        window.MouseUp(centre.Value, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static void ClickLink(Window window, string listPanelName, string content)
    {
        var button = window.FindControl<StackPanel>(listPanelName)!
            .GetVisualDescendants()
            .OfType<Button>()
            .First(b => (b.Content as string) == content);

        var centre = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window);
        Assert.NotNull(centre);

        window.MouseDown(centre!.Value, MouseButton.Left);
        window.MouseUp(centre.Value, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    // ------------------------------------------------------------- what the list holds

    // The completed list offers exactly the in-scope components - not everything the area touches.
    // C3 is inside the area but was never put in scope, so it is not work this entry tracks.
    [Fact]
    public void The_completed_list_offers_only_the_components_that_are_in_scope()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }), window =>
        {
            Assert.Equal(new[] { "C1", "C2" }, CompletedLabels(window));
        });
    }

    // A saved completed label comes back ticked, so reopening an entry shows the progress made
    // rather than resetting it.
    [Fact]
    public void Saved_progress_is_restored_when_the_entry_is_reopened()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2", "C3" }, completed: new[] { "C2" }), window =>
        {
            var byLabel = CompletedRows(window).ToDictionary(r => r.BoardLabel, r => r.IsChecked);

            Assert.False(byLabel["C1"]);
            Assert.True(byLabel["C2"]);
            Assert.False(byLabel["C3"]);
        });
    }

    // An entry with nothing in scope has nothing to track, so the list is empty and says so rather
    // than showing a bare empty box.
    [Fact]
    public void An_entry_with_nothing_in_scope_shows_the_empty_helper_text()
    {
        WithEditor(CreateEntry(Array.Empty<string>()), window =>
        {
            Assert.Empty(CompletedRows(window));
            Assert.True(window.FindControl<TextBlock>("EditorNoCompletedText")!.IsVisible);
        });
    }

    // ------------------------------------------------------------- staying in step with scope

    // THE rule that makes the list useful: newly-scoped work arrives as work still to do.
    [Fact]
    public void A_component_newly_ticked_into_scope_appears_unticked()
    {
        WithEditor(CreateEntry(new[] { "C1" }), window =>
        {
            Assert.Equal(new[] { "C1" }, CompletedLabels(window));

            ClickScopeRow(window, "C3");

            var byLabel = CompletedRows(window).ToDictionary(r => r.BoardLabel, r => r.IsChecked);

            Assert.True(byLabel.ContainsKey("C3"));
            Assert.False(byLabel["C3"]);
        });
    }

    [Fact]
    public void A_component_unticked_out_of_scope_leaves_the_completed_list()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }), window =>
        {
            ClickScopeRow(window, "C2");

            Assert.Equal(new[] { "C1" }, CompletedLabels(window));
        });
    }

    // Progress on OTHER components must survive a scope edit - the rows are rebuilt, so a naive
    // implementation would clear every tick each time the scope changed at all.
    [Fact]
    public void Progress_on_other_components_survives_a_scope_change()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }, completed: new[] { "C1" }), window =>
        {
            ClickScopeRow(window, "C3");

            var byLabel = CompletedRows(window).ToDictionary(r => r.BoardLabel, r => r.IsChecked);

            Assert.True(byLabel["C1"]);
        });
    }

    // Removing a component and putting it straight back starts it fresh, rather than restoring a
    // "done" flag for work that is now being tracked anew.
    [Fact]
    public void Removing_and_re_adding_a_component_forgets_that_it_was_completed()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }, completed: new[] { "C2" }), window =>
        {
            Assert.True(CompletedRows(window).Single(r => r.BoardLabel == "C2").IsChecked);

            ClickScopeRow(window, "C2");
            ClickScopeRow(window, "C2");

            Assert.False(CompletedRows(window).Single(r => r.BoardLabel == "C2").IsChecked);
        });
    }

    // Clearing the whole scope empties the completed list rather than leaving orphaned rows behind.
    [Fact]
    public void Clearing_the_scope_empties_the_completed_list()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2", "C3" }, completed: new[] { "C1" }), window =>
        {
            ClickLink(window, "EditorComponentScopePanel", "None");

            Assert.Empty(CompletedRows(window));
        });
    }

    [Fact]
    public void Selecting_all_in_scope_offers_every_component_as_outstanding()
    {
        WithEditor(CreateEntry(Array.Empty<string>()), window =>
        {
            ClickLink(window, "EditorComponentScopePanel", "All");

            Assert.Equal(new[] { "C1", "C2", "C3" }, CompletedLabels(window));
            Assert.All(CompletedRows(window), r => Assert.False(r.IsChecked));
        });
    }

    // ------------------------------------------------------------- the summary

    // The count reads as progress, since the list exists to answer "how much is left".
    [Fact]
    public void The_count_reports_progress_rather_than_a_bare_total()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2", "C3" }, completed: new[] { "C1", "C3" }), window =>
        {
            Assert.Equal("2 of 3 completed", window.FindControl<TextBlock>("EditorCompletedCountText")!.Text);
        });
    }

    [Fact]
    public void Marking_a_component_done_updates_the_count()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }), window =>
        {
            Assert.Equal("0 of 2 completed", window.FindControl<TextBlock>("EditorCompletedCountText")!.Text);

            ClickLink(window, "EditorComponentCompletedPanel", "All");

            Assert.Equal("2 of 2 completed", window.FindControl<TextBlock>("EditorCompletedCountText")!.Text);
        });
    }

    // ------------------------------------------------------------- dirty tracking

    // Building the list on open is not an edit - the same deferred-event trap the title box and the
    // scope checklist both fall into.
    [Fact]
    public void Populating_the_completed_list_does_not_mark_the_window_dirty()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }, completed: new[] { "C1" }), window =>
            Assert.False(window.FindControl<Button>("EditorSaveButton")!.IsEnabled));
    }

    [Fact]
    public void Marking_a_component_completed_marks_the_window_dirty()
    {
        WithEditor(CreateEntry(new[] { "C1", "C2" }), window =>
        {
            Assert.False(window.FindControl<Button>("EditorSaveButton")!.IsEnabled);

            ClickLink(window, "EditorComponentCompletedPanel", "All");

            Assert.True(window.FindControl<Button>("EditorSaveButton")!.IsEnabled);
        });
    }
}
