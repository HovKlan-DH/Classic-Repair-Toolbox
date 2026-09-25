using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The full editor's "Mark components in scope" checklist. The rows are supplied by the opener
// (TabSchematics owns the board data and highlight rectangles), so these tests drive the same
// public entry points the opener uses: Initialize, then InitializeComponentScope.
//
// The behaviour that matters, and is easy to get wrong:
//   - reopening an entry shows the selection the user saved, not everything re-ticked
//   - supplying a scope must not by itself mark the window dirty
//   - a window that was never given a scope must not wipe the entry's saved component list
[Collection("HeadlessUi")]
public class WorklogEntryScopeTests
{
    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static WorklogEntryRecord CreateEntry(params string[] componentLabels) => new()
    {
        Id = 1,
        SchematicName = "Sch",
        Title = "Bad cap",
        Category = "Issue",
        State = "Open",
        AreaX = 10,
        AreaY = 10,
        AreaWidth = 50,
        AreaHeight = 50,
        ComponentLabels = componentLabels.ToList(),
    };

    private static readonly (string BoardLabel, string DisplayName)[] Scope =
    {
        ("C12", "Capacitor | 100nF"),
        ("R4", "Resistor | 1k"),
        ("U1", "CPU | 6510"),
    };

    private static void WithEditor(WorklogEntryRecord entry, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            // Placement persistence off for this window only: it otherwise restores the size,
            // position and splitter ratio from the developer's REAL settings file, so every layout
            // assertion below would depend on how they last left the editor. Scoped rather than
            // assigned, so it cannot leak into other tests in the shared session.
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 900;
            window.Height = 700;

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

    private static IReadOnlyList<CheckBox> ScopeCheckBoxes(WorklogEntryEditorWindow window)
    {
        var list = window.GetVisualDescendants()
            .OfType<ItemsControl>()
            .First(c => c.Name == "EditorComponentList");

        return list.GetVisualDescendants().OfType<CheckBox>().ToList();
    }

    // Reopening an entry must show the choice the user made last time. Every row starting ticked
    // (which is what the quick "New fault" card does, since a new entry has no saved selection)
    // would silently re-add components the user had deliberately unticked the previous time.
    [Fact]
    public void Reopening_an_entry_ticks_only_the_components_it_had_saved()
    {
        bool[]? checkedStates = null;

        WithEditor(CreateEntry("C12", "U1"), window =>
        {
            window.InitializeComponentScope(Scope);
            Dispatcher.UIThread.RunJobs();

            checkedStates = ScopeCheckBoxes(window).Select(c => c.IsChecked == true).ToArray();
        });

        // C12 and U1 were saved on the entry; R4 was not.
        Assert.Equal(new[] { true, false, true }, checkedStates);
    }

    // Supplying the scope is not an edit - it is the window being populated. If it enabled Save,
    // every entry would look modified the instant it was opened.
    [Fact]
    public void Supplying_the_scope_does_not_enable_the_save_button()
    {
        bool? saveEnabled = null;

        WithEditor(CreateEntry("C12"), window =>
        {
            window.InitializeComponentScope(Scope);
            Dispatcher.UIThread.RunJobs();

            saveEnabled = window.GetVisualDescendants()
                .OfType<Button>()
                .First(b => b.Name == "EditorSaveButton")
                .IsEnabled;
        });

        Assert.False(saveEnabled);
    }

    // The checklist is hidden when no scope was supplied, rather than showing an empty box - an
    // empty list reads as "no components here", which is a different and wrong claim.
    [Fact]
    public void The_checklist_stays_hidden_when_no_scope_was_supplied()
    {
        bool? visible = null;

        WithEditor(CreateEntry("C12"), window =>
        {
            visible = window.GetVisualDescendants()
                .OfType<StackPanel>()
                .First(p => p.Name == "EditorComponentScopePanel")
                .IsVisible;
        });

        Assert.False(visible);
    }

    // An area that genuinely touches nothing still shows the section, with its "none" message -
    // that is a real answer, unlike the hidden case above.
    [Fact]
    public void An_empty_scope_still_shows_the_section_with_its_none_message()
    {
        bool? panelVisible = null;
        bool? noneVisible = null;

        WithEditor(CreateEntry(), window =>
        {
            window.InitializeComponentScope(Array.Empty<(string, string)>());
            Dispatcher.UIThread.RunJobs();

            panelVisible = window.GetVisualDescendants()
                .OfType<StackPanel>()
                .First(p => p.Name == "EditorComponentScopePanel")
                .IsVisible;

            noneVisible = window.GetVisualDescendants()
                .OfType<TextBlock>()
                .First(t => t.Name == "EditorNoComponentsText")
                .IsVisible;
        });

        Assert.True(panelVisible);
        Assert.True(noneVisible);
    }
}
