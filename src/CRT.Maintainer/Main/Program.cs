using System;
using Avalonia;
using Velopack;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // Entry point for the maintainer application (NewContributeStrategy.md Phase 5).
    //
    // *** VelopackApp.Build().Run() IS THE INSTALL HOOK, NOT AN UPDATER. *** The app is installed
    // by Velopack's Setup.exe, which starts the app with hook arguments while it installs,
    // updates and uninstalls, and this call answers them and exits before any window opens.
    // Without it "vpk pack" refuses to build the package at all ("Unable to verify VelopackApp is
    // called"), which is why no release of this app ever got past packing. Packing around that
    // check with --skipVeloAppCheck is not the fix: Setup.exe would then start the full app each
    // time it runs one of those hooks.
    //
    // CHECKING FOR UPDATES IS STILL DELIBERATELY ABSENT (no UpdateManager). This app's releases
    // sit in the same GitHub repository as CRT's, and Velopack's GithubSource merges the release
    // feeds of every recent release there without looking at the package id. The release
    // workflow keeps the two apart by publishing this app on its own channels (maintainer-win,
    // maintainer-linux - see build-and-release-maintainer.yml), and an updater added here must read
    // those, never CRT's "win"/"linux".
    // ###########################################################################################
    internal static class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            VelopackApp.Build().Run();
            BuildAvaloniaApp()
                .StartWithClassicDesktopLifetime(args);
        }

        // Referenced by name by Avalonia's XAML previewer, so it must stay public and keep this
        // exact signature even though nothing in this project calls it.
        public static AppBuilder BuildAvaloniaApp() =>
            AppBuilder.Configure<MaintainerApp>()
                .UsePlatformDetect()
                .WithInterFont()
                .LogToTrace();
    }
}
