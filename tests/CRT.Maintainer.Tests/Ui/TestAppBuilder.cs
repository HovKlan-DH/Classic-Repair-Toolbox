using Avalonia;
using Avalonia.Headless;

// The Avalonia application every headless UI test in this assembly runs against.
[assembly: AvaloniaTestApplication(typeof(CRT.Maintainer.Tests.Ui.TestAppBuilder))]

namespace CRT.Maintainer.Tests.Ui;

// ###########################################################################################
// The real maintainer application - its MaintainerApp.axaml with every theme dictionary and the
// table editor's colours - minus opening its main window. Its own start-up only sets the main
// window, which is exactly what a test must not have: MaintainerMain's OnOpened restores the
// signed-in session from the real AppData folder and fetches the queue from the live server. So
// no test ever Show()s a MaintainerMain; tests build one and read it.
// ###########################################################################################
public class HeadlessMaintainerApp : MaintainerApp
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
        return AppBuilder.Configure<HeadlessMaintainerApp>()
            .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = true });
    }
}
