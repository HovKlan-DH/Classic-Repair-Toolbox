using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;

namespace CRT.Review
{
    // ###########################################################################################
    // The application object (NewContributeStrategy.md Phase 5, task 1).
    //
    // KEPT DELIBERATELY EMPTY BEYOND SHOWING THE WINDOW. CRT's own App does a great deal here -
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
    public partial class ReviewApp : Application
    {
        public override void Initialize() => AvaloniaXamlLoader.Load(this);

        public override void OnFrameworkInitializationCompleted()
        {
            if (this.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                desktop.MainWindow = new ReviewMain();
            }

            base.OnFrameworkInitializationCompleted();
        }
    }
}
