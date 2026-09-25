using Avalonia;

namespace CRT;

// ###########################################################################################
// Test seams for the worklog area-marking flow.
//
// Same reasoning and same convention as TabSchematics.LabelEditor.TestSeams.cs: the three
// drawing methods are private and reached only through pointer handlers, so a headless test had
// no way to exercise the flow that "Add worklog" starts. Each member here delegates straight to
// the shipped method - nothing reimplements the behaviour.
//
// This flow is worth the seams: the parked-vs-anchored badge rule it feeds has been reported as
// a bug TWICE (once against the Workbooks board pane, once against the Schematics thumbnails),
// which makes it the flow in this tab most likely to regress unnoticed.
//
// WHERE THE COVERAGE STOPS. CompleteDrawingWorklogEntryRectangle ends by calling
// OpenNewWorklogEntryEditor, which needs a real owner Window for ShowDialog and so cannot run
// headlessly. That method was therefore split: TryFinishWorklogEntryDrawing holds the whole
// accept/reject decision, and both the shipped path and the seam below call it - so the rule
// under test is the one the application runs, with only the modal skipped.
//
// What these tests pin down is therefore the DRAWING and its accept/reject rule, NOT the editor
// handoff. Keep that distinction; a test claiming to cover the handoff would be lying.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    internal bool IsWorklogEntryModeForTests => this.thisIsWorklogEntryMode;

    internal bool IsDrawingWorklogEntryRectangleForTests => this.thisIsDrawingWorklogEntryRectangle;

    // The rubber-band rectangle shown while the drag is in progress; null when not drawing.
    internal Rect? WorklogEntryDraftRectangleForTests => this.thisWorklogEntryDraftRectangle;

    // The accepted area once the drag finishes; stays null when the drag was too small.
    internal Rect? WorklogEntryFinalRectangleForTests => this.thisWorklogEntryFinalRectangle;

    internal int WorklogEntryWorkbookIdForTests => this.thisWorklogEntryWorkbookId;

    internal void StartDrawingWorklogEntryRectangleForTests(Point startPixelPoint) =>
        this.StartDrawingWorklogEntryRectangle(startPixelPoint);

    internal void UpdateDrawingWorklogEntryRectangleForTests(Point currentPixelPoint) =>
        this.UpdateDrawingWorklogEntryRectangle(currentPixelPoint);

    // ###########################################################################################
    // Finishes the drag exactly as the shipped path does, stopping before the editor opens.
    //
    // This calls TryFinishWorklogEntryDrawing - the SAME method CompleteDrawingWorklogEntryRectangle
    // calls - so the accept/reject rule under test is the one the application runs, not a copy of
    // it. Only OpenNewWorklogEntryEditor is skipped, because ShowDialog needs a real owner Window.
    //
    // Returns true when the drag was accepted as a deliberate area.
    // ###########################################################################################
    internal bool CompleteWorklogEntryDrawingWithoutEditorForTests(Point releasePixelPoint) =>
        this.TryFinishWorklogEntryDrawing(releasePixelPoint, out _);
}
