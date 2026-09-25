using Avalonia;
using Avalonia.Headless;

// Registers the Avalonia application that every [AvaloniaFact] in this assembly runs
// against. Without this attribute the headless UI tests have no app to attach to.
[assembly: AvaloniaTestApplication(typeof(ClassicRepairToolbox.Tests.Ui.TestAppBuilder))]

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
