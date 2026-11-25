using System.Linq;
using Avalonia.Controls;
using Handlers.Geometry;

namespace CRT
{
    // ###########################################################################################
    // THE CATEGORY LIST ON THE FAR LEFT LEAVES THE COMPONENT LIST ITS HALF (2026-11-25, found looking for what else was
    // out of reach after the owner's report of the "Maintainer" tab's sign-in on Linux): the categories were as high as
    // their rows, and on CRT's smallest window the component list under them was 9 high. Each time the lists' grid
    // changes size the category lists - CRT's and the "Contribute" tab's own, one shown at a time - may be at most what
    // SideListsGeometry says; on any window with room that is more than their rows, and nothing changes.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        private void WireSideLists()
        {
            this.ComponentListsGrid.SizeChanged += (_, _) => this.FitCategoryLists();
        }

        private void FitCategoryLists()
        {
            // What the column leaves the two lists: its height less the words over each and the lists' own margins.
            double words = this.ComponentListsGrid.Children
                .Where(child => child.IsVisible && Grid.GetRow(child) is 0 or 2)
                .Sum(child => child.DesiredSize.Height);

            double margins = this.CategoryFilterListBox.Margin.Top + this.CategoryFilterListBox.Margin.Bottom
                + this.ComponentFilterListBox.Margin.Top + this.ComponentFilterListBox.Margin.Bottom;

            double most = SideListsGeometry.CategoryListMaxHeight(this.ComponentListsGrid.Bounds.Height - words - margins);

            foreach (ListBox list in new[] { this.CategoryFilterListBox, this.ContributeCategoryListBox })
            {
                if (!list.MaxHeight.Equals(most))
                {
                    list.MaxHeight = most;
                }
            }
        }
    }
}
