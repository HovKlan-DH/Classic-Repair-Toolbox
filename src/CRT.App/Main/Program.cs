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

            // The settings are written a moment after they change (UserSettings.Save): the last of
            // them go out here. ProcessExit does it as well, for the ways out that skip this.
            UserSettings.Flush();

            // The log is written in batches (Logger): the last lines of the session go out here,
            // before the process ends. ProcessExit does it as well, for the ways out that skip this.
            Logger.Flush();
        }

        // Avalonia configuration, don't remove; also used by visual designer.
        public static AppBuilder BuildAvaloniaApp()
            => AppBuilder.Configure<CRT.App>()
                .UsePlatformDetect()
                .With(Program.SkiaOptionsForCrt())
                .WithInterFont()
                .LogToTrace();

        // ###########################################################################################
        // *** ROOM ON THE GRAPHICS CARD FOR A SCHEMATIC IMAGE (2026-11-16, performance walkthrough). ***
        // Avalonia 12 keeps some 28 MB of pictures on the graphics card unless told otherwise, while one
        // schematic image of a large board is about 50 MB decoded - more than the whole room, so it was
        // suspected of being sent to the card again for every frame drawn while the picture is panned or
        // zoomed. AppConfig.GpuPictureBytes gives it room, the thumbnails and overlays beside it.
        // ###########################################################################################
        internal static SkiaOptions SkiaOptionsForCrt() => new() { MaxGpuResourceSizeBytes = CRT.AppConfig.GpuPictureBytes };
    }
}