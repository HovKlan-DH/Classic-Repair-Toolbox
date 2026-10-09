using Avalonia;
using Avalonia.Headless;

// Registers the Avalonia application that every [AvaloniaFact] in this assembly runs
// against. Without this attribute the headless UI tests have no app to attach to.
[assembly: AvaloniaTestApplication(typeof(ClassicRepairToolbox.Tests.Ui.TestAppBuilder))]

// ###########################################################################################
// *** THE APPLICATION IS BUILT ONCE FOR THE WHOLE ASSEMBLY, NOT PER TEST (2026-09-26). ***
//
// This is the fix for the random headless failure that blocked handovers three times in one day
// and had been seen since 2026-09-21: a test failing in 1-3 ms with
//
//     InvalidOperationException: The calling thread cannot access this object because a
//     different thread owns it
//       at Dispatcher.VerifyAccess <- DefaultRenderLoop.Add <- ServerCompositor..ctor
//       <- AvaloniaHeadlessPlatform.Initialize <- HeadlessUnitTestSession.EnsureIsolatedApplication
//
// - i.e. inside Avalonia's own per-test setup, BEFORE the named test's body runs, which is why it
// landed on a different innocent test each time (NewBoardWindowTests, OscilloscopeSequencingTests,
// DraftedHighlightMarkingTests) and why each passed in isolation.
//
// Avalonia's DEFAULT is AvaloniaTestIsolationLevel.PerTest, which tears down and rebuilds the
// Application AND ITS DISPATCHER before every single test. That rebuild is what throws: the
// previous test's dispatcher work is not always finished settling when the new dispatcher is
// constructed, and the new render loop then verifies thread affinity against the old one.
//
// PerAssembly builds it once and is what this suite has always ASSUMED - UiTest holds one session
// per assembly for exactly that reason, and every UI test is already in the one "HeadlessUi"
// collection sharing that single dispatcher thread, so no test here relies on a fresh Application.
// The tests that touch global state (UserSettings, DataManager, WorklogManager) isolate themselves
// through their own seams, never through Avalonia's teardown.
// ###########################################################################################
[assembly: AvaloniaTestIsolation(AvaloniaTestIsolationLevel.PerAssembly)]

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The real CRT.App, minus its startup sequence.
//
// Inheriting from it means the tests get the genuine App.axaml - every theme dictionary,
// brush and style the tabs actually bind to - so a resource key deleted from App.axaml is
// caught here rather than at runtime on a user's machine.
//
// OnFrameworkInitializationCompleted is deliberately NOT called through to base. The real
// one calls Logger.Initialize() (which .claude/CLAUDE.md forbids tests from doing, because
// it writes to the user's real log file), shows a splash screen, syncs data over the
// network and opens the main window. None of that belongs in a test run.
// ###########################################################################################
public class HeadlessTestApp : CRT.App
{
    public override void OnFrameworkInitializationCompleted()
    {
        // Intentionally empty - see the class comment above.
    }
}

public static class TestAppBuilder
{
    public static AppBuilder BuildAvaloniaApp()
    {
        return AppBuilder.Configure<HeadlessTestApp>()
            .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = true });
    }
}
