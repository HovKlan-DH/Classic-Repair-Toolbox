using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE MAINTAINER VIEW'S LIST: who maintains the system, one line each - and nothing more (owner
    // request, 2026-10-04: "please make sure that the actual 'Send invitation' and 'Add as
    // maintainer' gets moved to the 'Admin' tab (in top menu), as this is something only the admin
    // should be able to do. The stuff that should be visible in here, is just the selected
    // maintainer(s)").
    //
    // Adding, inviting, removing and withdrawing are the administrator's, under Account > Maintainers
    // (MaintainerPoolView) - here from 2026-09-27 to 2026-10-04. Every maintainer sees the same list,
    // with the addresses only on a system they maintain themselves (2026-10-05, SystemOverviewFlow).
    // ###########################################################################################
    public partial class SystemView
    {
        private void ShowMaintainers(SystemDetailAnswer? detail)
        {
            if (this.FindControl<StackPanel>("MaintainersSection") is not StackPanel section)
                return;

            section.Children.Clear();

            if (detail is null)
                return;

            section.Children.Add(SystemView.Heading(SystemsDisplay.MaintainersHeading(detail.Maintainers.Count), detail.Maintainers.Count == 0));

            // Names only for a system this account does not maintain (2026-10-05) - said, once.
            if (detail.AddressesHidden && detail.Maintainers.Count > 0)
                section.Children.Add(SystemView.Note(SystemsDisplay.AddressesHiddenLine));

            foreach (PoolMaintainerEntry maintainer in detail.Maintainers)
                section.Children.Add(SystemView.Line(SystemsDisplay.MaintainerLine(maintainer)));
        }
    }
}
