using Handlers.DataHandling;
using System;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // WHEN A BOARD COUNTS AS VIEWED (owner request, 2026-09-27): "I want it to catch every time a
    // user selects a board - not only once per day, but it is a good idea to have the count only
    // when viewed for +10 seconds."
    //
    // One view per time a board COMES ON SCREEN and stays there for BoardViewRules.
    // MinimumTimeOnScreen. Main tells this which board is on screen after every board load
    // (Show), and counts when TakeDue answers:
    //
    //   - Choosing another board starts a new view; one left within the ten seconds is not counted,
    //     so flicking through the drop-down counts nothing.
    //   - The SAME board shown again is not a new view - a load that re-reads the board on screen
    //     (hiding a schematic in the Configuration tab, a draft saved) is not the user choosing it.
    //     Going to another board and back is: A, B, A counts A twice.
    //   - The board restored at launch is a view like any other: it is on screen.
    //   - A view is counted once, however long the board then stays.
    //
    // Pure - the clock is handed in - so the rule is tested; the timer is Main.BoardViews'.
    // ###########################################################################################
    public sealed class BoardViewTracker
    {
        private string? thisShown;
        private DateTimeOffset thisSince;
        private bool thisCounted;

        // The board on screen (its board id), or null.
        public string? Shown => this.thisShown;

        // ###########################################################################################
        // The board on screen is now `boardId` (null or blank: none, or one that is not counted).
        // True when that CHANGED the board on screen - a view of the new one starts now, and any
        // view of the old one still waiting to count never will.
        // ###########################################################################################
        public bool Show(string? boardId, DateTimeOffset now)
        {
            string? shown = string.IsNullOrWhiteSpace(boardId) ? null : boardId.Trim();

            if (string.Equals(shown, this.thisShown, StringComparison.Ordinal))
                return false;

            this.thisShown = shown;
            this.thisSince = now;
            this.thisCounted = false;
            return true;
        }

        // How long until the board on screen counts (zero: now); null when nothing is waiting to.
        public TimeSpan? Remaining(DateTimeOffset now)
        {
            if (this.thisShown is null || this.thisCounted)
                return null;

            TimeSpan left = BoardViewRules.MinimumTimeOnScreen - (now - this.thisSince);

            return left > TimeSpan.Zero ? left : TimeSpan.Zero;
        }

        // The board to count - once - when it has been on screen long enough; null otherwise.
        public string? TakeDue(DateTimeOffset now)
        {
            if (this.Remaining(now) != TimeSpan.Zero)
                return null;

            this.thisCounted = true;
            return this.thisShown;
        }
    }
}
