using Avalonia.Controls;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Stands in for CRT's window around the Maintainer tab, as far as waiting goes (2026-09-29).
//
// The tab has no "please wait" overlay of its own - every wait in it runs under CRT's window's one
// (Main.axaml), which BusyOverlay.For finds by walking up to the root and then down. Main is too
// much to build for a test about a wait, so this puts the tab and one BusyOverlay in the same
// Grid, the overlay LAST - Main.axaml's own arrangement. The tests that used to read the old
// window's overlay by name read this one instead.
// ###########################################################################################
internal static class MaintainerTabHost
{
    // The overlay a wait anywhere in `tab` now runs under. `tab` must not have a parent yet.
    internal static BusyOverlay AddOverlay(TabMaintainer tab)
    {
        var overlay = new BusyOverlay();

        _ = new Grid { Children = { tab, overlay } };

        return overlay;
    }
}
