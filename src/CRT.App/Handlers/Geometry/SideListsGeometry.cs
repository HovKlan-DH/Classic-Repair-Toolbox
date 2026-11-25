using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE CATEGORY AND COMPONENT LISTS ON THE FAR LEFT SHARE THE ROOM UNDER "Board" (2026-11-25, found looking for what
    // else was out of reach after the owner's report of the "Maintainer" tab's sign-in on Linux, "it makes the UI useless
    // in these situation"): the category list was as high as its rows, the component list had what was left - and on
    // CRT's smallest window the C64 250407's seven categories took all of it, the component list 9 high, not one row to
    // pick.
    //
    // The categories keep every row while that leaves the components at least half the room; else they get half and
    // scroll. On any window with room nothing changes. Never lower than two rows.
    // ###########################################################################################
    public static class SideListsGeometry
    {
        // Two of the category list's rows.
        public const double CategoryListLeast = 44;

        // The tallest the category list may be, the two lists sharing `room` - what their column leaves them, less the
        // words over each. A room not laid out yet (none) limits nothing.
        public static double CategoryListMaxHeight(double room)
        {
            if (!(room > 0) || double.IsInfinity(room))
            {
                return double.PositiveInfinity;
            }

            return Math.Max(SideListsGeometry.CategoryListLeast, room / 2);
        }
    }
}
