using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How Account > "Order of systems" reads (owner request, 2026-10-04: "In the 'Admin' top menu I
    // would like a possibility to be able to sort the list of systems, which then gets saved to both
    // sources (BETA + stable) after my save").
    //
    // The administrator drags BETA's drop-down list into order and saves; the server writes that
    // order into BETA's and the stable source's main Excel data files (CRT.Server's SystemOrderFlow).
    // Pure, so the words and the "has it moved" rule are tested; SystemOrderView only draws.
    // ###########################################################################################
    public static class SystemOrderDisplay
    {
        public const string Heading = "Order of systems";

        public const string Explanation =
            "The order CRT shows the hardware and boards in its drop-down lists. Drag a system up or down, then save - " +
            "the order is written into the main Excel data file of BETA and of the stable source, and CRT downloads it " +
            "on its next launch.";

        public const string NoList =
            "BETA has no versioned main Excel data file, so there is no list to put in order.";

        public const string SaveButton = "Save the order";

        public const string UndoButton = "Cancel";

        // Under the list while it has been moved and not saved - the one thing that says Save matters.
        public const string Unsaved = "The order has changed and is not saved yet.";

        // ###########################################################################################
        // Whether the list on screen is in another order than the one read - by system, so renaming
        // nothing and moving nothing is never a change.
        // ###########################################################################################
        public static bool IsReordered(IReadOnlyList<string> asRead, IReadOnlyList<string> onScreen)
        {
            ArgumentNullException.ThrowIfNull(asRead);
            ArgumentNullException.ThrowIfNull(onScreen);

            return !asRead.SequenceEqual(onScreen, StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The ids a save sends: each system ONCE, where it first appears (any case). A main Excel data
        // file can list one system on two rows - two workbooks in one system folder - and the server
        // refuses a list naming a system twice, so such a list could never be saved (code review,
        // 2026-10-04). Once is all the server needs: its MasterListing.ArrangeAs keeps a system's
        // rows together, in their own order, at the place the system's first id gives.
        // ###########################################################################################
        public static IReadOnlyList<string> OrderToSend(IEnumerable<string> onScreen)
        {
            ArgumentNullException.ThrowIfNull(onScreen);

            var sent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            return onScreen.Where(sent.Add).ToList();
        }

        // ###########################################################################################
        // What a save did, and whether that is all good. BETA first: it is the list dragged. The
        // stable source is named only when the server has one; a list already in this order says
        // so rather than claiming a write. Anything the server could not do after writing BETA's
        // list (the stable list, a manifest) follows, and makes the line an error.
        // ###########################################################################################
        public static (string Text, bool IsError) Saved(SystemOrderAnswer answer)
        {
            ArgumentNullException.ThrowIfNull(answer);

            string? problem = string.IsNullOrWhiteSpace(answer.Problem) ? null : answer.Problem.Trim();
            string text;

            if (!answer.BetaChanged && answer.StableChanged != true)
            {
                text = answer.StableChanged is null
                    ? "Nothing to save - BETA's list was already in this order."
                    : problem is null
                        ? "Nothing to save - the lists in BETA and the stable source were already in this order."
                        : "BETA's list was already in this order.";
            }
            else
            {
                text = answer.StableChanged switch
                {
                    null => "Saved. CRT's drop-down list in BETA is in the new order (this server has no stable source).",
                    true when answer.BetaChanged => "Saved. CRT's drop-down lists in BETA and the stable source are in the new order.",
                    true => "Saved. BETA's list was already in this order; the stable source's list is now in it too.",

                    // Not changed: already in this order - unless the problem below says it could not be.
                    false when problem is null => "Saved. CRT's drop-down list in BETA is in the new order; the stable source's was already in it.",
                    false => "Saved. CRT's drop-down list in BETA is in the new order."
                };
            }

            return problem is null ? (text, false) : ($"{text} But: {problem}", true);
        }
    }
}
