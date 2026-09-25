using System;
using Avalonia;

namespace CRT.Review
{
    // ###########################################################################################
    // Entry point for the review application (NewContributeStrategy.md Phase 5).
    //
    // DELIBERATELY THINNER THAN CRT'S OWN Program.cs, which initialises Velopack before Avalonia
    // because the app updates itself in a user's hands. This app ships through its own workflow
    // (task 8) and updating is added when that workflow exists - not stubbed here, where a
    // half-wired updater would be worse than none.
    // ###########################################################################################
    internal static class Program
    {
        [STAThread]
        public static void Main(string[] args) => Program
            .BuildAvaloniaApp()
            .StartWithClassicDesktopLifetime(args);

        // Referenced by name by Avalonia's XAML previewer, so it must stay public and keep this
        // exact signature even though nothing in this project calls it.
        public static AppBuilder BuildAvaloniaApp() =>
            AppBuilder.Configure<ReviewApp>()
                .UsePlatformDetect()
                .WithInterFont()
                .LogToTrace();
    }
}
