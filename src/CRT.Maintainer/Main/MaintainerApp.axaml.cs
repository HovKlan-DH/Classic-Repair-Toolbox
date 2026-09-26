using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using CRT.Maintainer.Handlers;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // The application object (NewContributeStrategy.md Phase 5, task 1).
    //
    // KEPT DELIBERATELY EMPTY BEYOND SHOWING THE WINDOW (and handing it its remembered settings,
    // 2026-09-26). CRT's own App does a great deal here -
    // theme JSON, global exception logging, a splash, a data sync - and every one of those is a
    // thing this app either does not have yet or should not copy on the assumption it will. They
    // are added when something needs them.
    //
    // *** ANYTHING ADDED HERE MUST NOT BREAK HEADLESS TESTING. *** CRT learned this the hard way:
    // its OnFrameworkInitializationCompleted calls Logger.Initialize(), shows a splash and syncs
    // over the network, so the headless test harness has to subclass it with an EMPTY override
    // (see CRT.App.Tests' TestAppBuilder). Keeping startup work out of here means this app's own
    // tests never need that workaround.
    // ###########################################################################################
    public partial class MaintainerApp : Application
    {
        public override void Initialize() => AvaloniaXamlLoader.Load(this);

        public override void OnFrameworkInitializationCompleted()
        {
            if (this.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                // Where the window was and "Show changes only", from the last run - read and
                // written here only, so a test building the window never touches the real file.
                string settingsPath = MaintainerSettingsStore.DefaultPath;

                var main = new MaintainerMain();
                main.UseSettings(
                    MaintainerSettingsStore.Load(settingsPath),
                    settings => MaintainerSettingsStore.Save(settingsPath, settings));

                desktop.MainWindow = main;
            }

            base.OnFrameworkInitializationCompleted();
        }
    }
}
