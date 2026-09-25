using Velopack;
using Avalonia;
using Handlers.DataHandling;
using System;

namespace Main
{
    internal class Program
    {
        // Initialization code. Don't use any Avalonia, third-party APIs or any
        // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
        // yet and stuff might break.
        [STAThread]
        public static void Main(string[] args)
        {
            // First thing that happens, because this is the only point where the raw command line
            // is available before anything can read SimulationOptions.Current. Pure parsing - no
            // Avalonia, no Logger (which does not exist yet); the result is logged from App once
            // the log file is open.
            SimulationOptions.Initialize(args);

            VelopackApp.Build().Run();
            BuildAvaloniaApp()
                .StartWithClassicDesktopLifetime(args);
        }

        // Avalonia configuration, don't remove; also used by visual designer.
        public static AppBuilder BuildAvaloniaApp()
            => AppBuilder.Configure<CRT.App>()
                .UsePlatformDetect()
                .WithInterFont()
                .LogToTrace();
    }
}