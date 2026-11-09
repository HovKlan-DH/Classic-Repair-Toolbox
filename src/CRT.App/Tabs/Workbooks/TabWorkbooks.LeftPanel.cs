using Avalonia;
using Avalonia.Controls;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // WHAT OF THIS TAB STANDS IN MAIN'S LEFT PANEL: "My devices" while the tab is shown, and the workbook list in its place
    // while no device is made. In a part of its own since 2026-11-09 (code review: TabWorkbooks.axaml.cs had grown past
    // 1,500 lines).
    //
    // Part of the TabWorkbooks partial class - see TabWorkbooks.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // ###########################################################################################
        // "MY DEVICES" STANDS IN MAIN'S LEFT PANEL (owner requests, 2026-11-06: "when in the 'My Devices' tab, I want
        // the 'My devices' list ... to show in the left-side panel, where the component category and component list
        // is", then "please again have [the workbooks] visible where they were previously, but do keep the 'My
        // devices' in the left-side panel"). Main hands over the place, once, at construction (Main.DeviceList.cs),
        // and shows it only while this tab is; the line between it and the workbooks goes with it. Everything else
        // about the list is as it was - its controls are this tab's, filled by this tab's code, named in this tab's
        // name scope - and its card styles are its own (TabWorkbooksCardStyles.axaml).
        //
        // Its heading over it as "Hardware" and "Board" are over theirs, and the list framed as the left panel's
        // lists are, a line all round, the cards and "Clear" a little in from it.
        // ###########################################################################################
        internal void ShowDevicesIn(Decorator place)
        {
            if (this.DeviceListArea.Parent is Panel column)
            {
                column.Children.Remove(this.DeviceListArea);
            }

            // The workbook list's line, kept to the theme as its DynamicResource is (code review, 2026-11-07: copied once,
            // it stayed the old theme's color after a switch of theme).
            this.DeviceListBorder.Bind(Border.BorderBrushProperty, this.DeviceListBorder.GetResourceObservable("Workbooks_Separator"));
            this.DeviceListBorder.BorderThickness = new Thickness(1);
            this.DeviceListBorder.CornerRadius = new CornerRadius(4);
            this.DeviceListBorder.ClipToBounds = true;

            if (this.DeviceListBorder.Child is Control inside)
            {
                inside.Margin = new Thickness(5);
            }

            this.DeviceListSeparator.IsVisible = false;
            place.Child = this.DeviceListArea;
            this.thisLeftPanelPlace = place;
        }

        // ###########################################################################################
        // *** WITH NO DEVICE MADE, THE WORKBOOKS STAND IN THE LEFT PANEL IN PLACE OF "My devices" (owner request,
        // 2026-11-08: "when there are no devices created, then the list of workbooks should move into the left-side
        // panel, as there are no devices to display there. This would make the content area larger for the actual
        // worklogs"). *** The whole list moves - its heading and count over a frame, as "My devices" stands there, and
        // "Clear" under it - and the tab's own list column and splitter close, the workbook taking the tab's width. The
        // first device made puts both back (RefreshDeviceList, at every refresh). The list's controls stay this tab's.
        // A tab Main has handed no place keeps the list where it is.
        // ###########################################################################################
        private Decorator? thisLeftPanelPlace;

        private bool thisWorkbooksInLeftPanel;

        internal bool IsWorkbookListInLeftPanel => this.thisWorkbooksInLeftPanel;

        private void PlaceWorkbookList(bool noDevices)
        {
            if (this.thisLeftPanelPlace is not { } place || noDevices == this.thisWorkbooksInLeftPanel)
            {
                return;
            }

            this.thisWorkbooksInLeftPanel = noDevices;

            if (noDevices)
            {
                this.WorkbookListBorder.Child = null;
                place.Child = this.WorkbookListArea;

                // The frame "My devices" has there (ShowDevicesIn).
                this.WorkbookListFrame.Bind(Border.BorderBrushProperty, this.WorkbookListFrame.GetResourceObservable("Workbooks_Separator"));
                this.WorkbookListFrame.Bind(Border.BackgroundProperty, this.WorkbookListFrame.GetResourceObservable("Workbooks_Panel_Bg"));
                this.WorkbookListFrame.BorderThickness = new Thickness(1);
                this.WorkbookListFrame.CornerRadius = new CornerRadius(4);
                this.WorkbookListFrame.ClipToBounds = true;
                this.WorkbookListFrame.Margin = new Thickness(0, 8, 0, 0);
                this.WorkbookListScroller.Margin = new Thickness(5);
                this.ClearWorkbookButton.Margin = new Thickness(0, 6, 0, 0);
            }
            else
            {
                place.Child = this.DeviceListArea;
                this.WorkbookListBorder.Child = this.WorkbookListArea;

                this.WorkbookListFrame.ClearValue(Border.BorderBrushProperty);
                this.WorkbookListFrame.ClearValue(Border.BackgroundProperty);
                this.WorkbookListFrame.BorderThickness = default;
                this.WorkbookListFrame.CornerRadius = default;
                this.WorkbookListFrame.ClipToBounds = false;
                this.WorkbookListFrame.Margin = default;
                this.WorkbookListScroller.Margin = new Thickness(9, 9, 9, 0);
                this.ClearWorkbookButton.Margin = new Thickness(9, 6, 9, 15);
            }

            // Its heading and count badge there, the count line in the tab.
            this.WorkbookListHeading.IsVisible = noDevices;
            this.WorkbookCountText.IsVisible = !noDevices;

            // The tab's own list column: shut while the list is away, at its width again when it is back.
            this.WorkbookListBorder.IsVisible = !noDevices;
            this.OuterSplitter.IsVisible = !noDevices;

            ColumnDefinition list = this.OuterSplitGrid.ColumnDefinitions[0];
            list.MinWidth = noDevices ? 0 : MinimumPanelWidth;
            list.Width = new GridLength(noDevices ? 0 : ClampPanelWidth(UserSettings.WorkbooksLeftPanelWidth));
            this.OuterSplitGrid.ColumnDefinitions[1].Width = new GridLength(noDevices ? 0 : 4);
        }
    }
}
