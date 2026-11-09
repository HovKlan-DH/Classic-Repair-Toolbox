namespace CRT
{
    // ###########################################################################################
    // "MY DEVICES" IN THE LEFT PANEL WHILE THE "Workbooks" TAB IS SHOWN (owner requests, 2026-11-06: "when in the
    // 'My Devices' tab, I want the 'My devices' list ... to show in the left-side panel, where the component
    // category and component list is. When in this tab, I have no use of those data"; then "please again have [the
    // workbooks] visible where they were previously, but do keep the 'My devices' in the left-side panel").
    //
    // The tab hands its list of devices over once, at construction (TabWorkbooks.ShowDevicesIn) - it stays the
    // tab's, filled by the tab's own code; the workbooks stay in the tab. While the tab is shown, the category and
    // component lists and the buttons under them give way to it; Hardware and Board stay. The Contribute tab's own
    // lists stand in the component lists' place (Main.ContributeLists.cs), so this hides and shows their container
    // whole and never a list in it, which that tab switches.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // Once, in the constructor: the tab's list of devices into its place in the left panel.
        private void PlaceDeviceList()
        {
            this.TabWorkbooks.ShowDevicesIn(this.DeviceListHost);
        }

        // At every switch of tab: the devices while the tab is shown, the component lists otherwise.
        private void FollowWorkbooksTabInLeftPanel()
        {
            bool shown = this.WorkbooksTabItem is not null && ReferenceEquals(this.MainTabControl?.SelectedItem, this.WorkbooksTabItem);

            this.ComponentListsGrid.IsVisible = !shown;
            this.ComponentActionsPanel.IsVisible = !shown;
            this.DeviceListHost.IsVisible = shown;
        }
    }
}
