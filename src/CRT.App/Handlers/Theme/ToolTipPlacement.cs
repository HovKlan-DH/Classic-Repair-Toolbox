using Avalonia.Controls;
using Avalonia.Data;

namespace Handlers.Theming
{
    // ###########################################################################################
    // WHERE EVERY TOOLTIP IN CRT OPENS: just below the control it belongs to - or just above it
    // where there is no room below - never under the pointer (owner report, 2026-09-30: "I need to
    // click the "Approve and publish to BETA" multiple times for it to react").
    //
    // *** AVALONIA'S DEFAULT PUT A TOOLTIP OVER THE POINTER AT THE BOTTOM OF THE SCREEN. *** It opens
    // at the POINTER, 20 px below it. With no room below, the positioner flips it above the pointer -
    // and keeps the +20 offset, which moves the flipped tooltip back down over the pointer. CRT's
    // tooltips are native popup windows, so the next click went to the tooltip rather than to the
    // control under it; the tooltip closed, opened again after a moment's hover, and a button read
    // as needing two or three clicks. Anything in the bottom hundred pixels or so of a maximised
    // window was exposed: the decision buttons, and every row of a table, list or file tree that
    // runs to the bottom (the table's cell tooltips open at once).
    //
    // *** SO THE DEFAULT IS CHANGED FOR EVERY CONTROL, ONCE (App's static constructor). *** Placed
    // at the control's own bottom edge with no offset, a tooltip either sits under the control or -
    // flipped - on top of it, touching it and never covering it, so the pointer on the control is
    // never under it. One default rather than a setting on seventy tooltips, so a tooltip added
    // tomorrow is safe without anybody remembering. A control that sets a placement of its own
    // still wins.
    //
    // ToolTipPlacementTests reproduces the lost click and fails against Avalonia's default.
    // ###########################################################################################
    public static class ToolTipPlacement
    {
        // ###########################################################################################
        // *** AS EACH TOOLTIP OPENS, NOT AS A DEFAULT. *** Avalonia registers these two attached
        // properties' defaults on Control itself, so they cannot be overridden there ("Metadata is
        // already set for Placement on Control") - and overriding them type by type would miss
        // every type not listed. A class handler on ToolTipOpening sees every control whose tooltip
        // is about to open, and sets the two before Avalonia places it.
        // ###########################################################################################
        public static void UseControlEdgePlacement()
        {
            ToolTip.ToolTipOpeningEvent.AddClassHandler<Control>(
                (control, _) => ToolTipPlacement.PlaceAtEdge(control),
                handledEventsToo: true);
        }

        // ###########################################################################################
        // Below the control, or above it where there is no room - unless the control chose its own.
        //
        // *** WRITTEN AT STYLE PRIORITY, NOT AS A LOCAL VALUE (code review, 2026-10-01). *** This is
        // a DEFAULT, and it is written onto the control the first time its tooltip opens. As a local
        // value it was indistinguishable from a choice the control's author made, and it outranked
        // every style and trigger set afterwards - a placement applied by a style on a later theme
        // change lost to it for good. At Style priority it is only ever the weakest opinion: a local
        // value, a style trigger or anything else the author sets wins, whenever it is set.
        // ###########################################################################################
        public static void PlaceAtEdge(Control control)
        {
            System.ArgumentNullException.ThrowIfNull(control);

            if (!control.IsSet(ToolTip.PlacementProperty))
                control.SetValue(ToolTip.PlacementProperty, PlacementMode.Bottom, BindingPriority.Style);

            if (!control.IsSet(ToolTip.VerticalOffsetProperty))
                control.SetValue(ToolTip.VerticalOffsetProperty, 0d, BindingPriority.Style);
        }
    }
}
