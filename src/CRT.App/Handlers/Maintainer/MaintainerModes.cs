using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.Theming;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // THE FOUR SCREENS BEHIND THE BUTTONS AT THE TOP LEFT (owner request, 2026-09-27):
    //
    //   Review   - the queue of contributions, as it always was.
    //   BETA     - boards in BETA waiting to be published to production (the "Production" window).
    //   Boards  - every board: its maintainers, contributors and submissions.
    //   Account  - "My account" and "Server version" for every maintainer, and the administrator's
    //              own entries below them (it was "Admin", the administrator's alone, until
    //              2026-10-04).
    //
    // Each shows its list on the left and what is chosen in it on the right, where three separate
    // windows used to open over the queue.
    //
    // *** THE BADGES COUNT BOARDS, NOT SUBMISSIONS ***, because that is what was asked for ("how
    // many boards need your attention") and what a maintainer works through: two submissions to one
    // board are reviewed together. "Needs your attention" is the SERVER's answer each time - the
    // queue's AwaitsYou, the production list's AwaitsYou - so a board waiting only for the other
    // approver is not counted. Pure, so the counting is tested.
    // ###########################################################################################
    public enum MaintainerMode
    {
        Review,
        Beta,
        Boards,
        Account
    }

    public static class MaintainerModes
    {
        // The most a badge spells out; more reads as "99+" - TabBadge's rule, which the tab badges
        // in CRT's row of tabs share.
        public const int BadgeCeiling = TabBadge.Ceiling;

        // ###########################################################################################
        // Review: the boards with at least one submission waiting for THIS account. A submission the
        // server does not say about (an older server, null) counts as yours - the queue shows it as
        // yours too, undimmed.
        // ###########################################################################################
        public static int ReviewAttention(IEnumerable<ReviewQueueRow>? rows) =>
            rows is null
                ? 0
                : rows.Where(row => row.AwaitsYou != false)
                    .Select(row => row.BoardId ?? string.Empty)
                    .Distinct(StringComparer.Ordinal)
                    .Count();

        // BETA: the boards waiting for production that wait for THIS account - not the ones this
        // account has already approved, which wait for the other approver.
        public static int BetaAttention(IEnumerable<ProductionBoardRow>? rows) =>
            rows?.Count(row => row.AwaitsYou != false) ?? 0;

        // ###########################################################################################
        // THE MAINTAINER TAB'S OWN BADGE, in CRT's row of tabs (owner request, 2026-09-30: "when a
        // maintainer is logged in, and he/she receives something new in queue (for him to process),
        // then it should show as a badge in the "Maintainer" tab ... until it is fully processed,
        // including if it is awaiting in "BETA to PROD" queue").
        //
        // *** THE SUM OF THE TWO BUTTONS' BADGES, NOT A COUNT OF ITS OWN. *** A board with a new
        // submission queued AND an earlier one in BETA waits for this account twice, and each button
        // counts it once - so the tab says what opening it will show: "3" over a "2" and a "1".
        // Counting distinct boards would say 2 and leave the maintainer adding the buttons up to 3.
        // Nothing waiting for this account in either - including a submission waiting only for the
        // other approver - is no badge.
        // ###########################################################################################
        public static int TabAttention(IEnumerable<ReviewQueueRow>? queue, IEnumerable<ProductionBoardRow>? beta) =>
            MaintainerModes.ReviewAttention(queue) + MaintainerModes.BetaAttention(beta);

        // ###########################################################################################
        // WHETHER THE TAB IS SHOWN AT ALL (owner request, 2026-10-01: "Show only Maintainer tab when
        // I have outstanding work ... If NOT checked, then it will not show when there is no badge
        // on the "Maintainer" tab, as that mean there is no work for the maintainer to do").
        //
        // Three inputs, in the order they decide it:
        //
        //   enabled         - "Enable Maintainer tab". Off hides the tab whatever else is true; this
        //                     rule narrows that setting, it never overrides it.
        //   onlyWhenWaiting - the new checkbox. Off means the tab is simply shown, as before.
        //   badgeKnown      - the number is a real answer (see below); an unread badge never hides the tab.
        //   attention       - TabAttention, the badge's own number. Nothing waiting, no tab.
        //   selected, unsavedEdits - hold a tab open whatever the number says; see below.
        //
        // *** A TAB THE MAINTAINER IS STANDING ON IS NEVER TAKEN AWAY (owner decision, 2026-10-01).
        // *** `selected` holds it open. Approving the last submission in the queue drops the badge
        // to zero the moment it is decided - and without this the tab, the submission's table and
        // any unsaved edits in it would vanish from under the maintainer at exactly that moment.
        // It goes on the next tab switch instead, which is when the maintainer has finished with it.
        // Unticking "Enable Maintainer tab" still hides it at once; that is a deliberate act.
        //
        // *** NOR ONE HOLDING UNSAVED TABLE EDITS (code review, 2026-10-01). *** `selected` alone let
        // it go: a maintainer edits the table of a submission that waits only for the OTHER
        // approver (so it counts nothing), switches to another tab without saving, and the next
        // minute check hid the tab with the edits inside it - nothing would bring it back short of
        // the setting, and quitting CRT then selected an invisible tab to ask about them.
        // `unsavedEdits` holds it open until they are saved or discarded.
        //
        // *** `badgeKnown`: THE BADGE'S NUMBER IS A REAL ANSWER, NOT JUST ZERO. *** The count is zero
        // before anything was read, after a 401 empties it (ShowSignInPanel), and while the server
        // cannot be reached - none of which is "no work". TabMaintainer.BadgeKnown says the account
        // is signed in and the queue and the BETA list have each been read since (code review,
        // 2026-10-01: it used to be "signed in", which a restored session satisfied before any list
        // arrived). Only then does zero hide the tab; otherwise it stays, so a maintainer can always
        // get back to the sign-in screen and a failed read never makes the tab vanish.
        // ###########################################################################################
        public static bool TabIsShown(bool enabled, bool onlyWhenWaiting, bool badgeKnown, int attention, bool selected, bool unsavedEdits)
        {
            if (!enabled)
                return false;

            if (!onlyWhenWaiting || !badgeKnown || selected || unsavedEdits)
                return true;

            return attention > 0;
        }

        // ###########################################################################################
        // WHICH ENTRY "CONTRIBUTOR SUBMISSIONS" AND "BETA > PROD" OPEN ON (owner request, 2026-09-30:
        // "it should select either whatever the user viewed last time (if set) or it should show the
        // first entry, instead of the maintainer needing to press it"): the one looked at last, while
        // it is still in the list - else the first. Null only for an empty list. `inListOrder` is
        // the keys as the list shows them, top to bottom.
        // ###########################################################################################
        public static long? EntryToOpen(IReadOnlyList<long> inListOrder, long? remembered)
        {
            ArgumentNullException.ThrowIfNull(inListOrder);

            if (remembered is long id && inListOrder.Contains(id))
                return id;

            return inListOrder.Count > 0 ? inListOrder[0] : null;
        }

        // ###########################################################################################
        // THE SCREEN THE TAB OPENS ON (owner request, 2026-10-04: "if there is no queue awaiting,
        // when opening the "Maintainer" tab, then go to "Systems" and show the last selected system").
        //
        // `shown` is the screen on show as the tab is opened (Review after signing in), `attention`
        // the tab's badge (TabAttention) and `somethingOpen` whether something is open on that screen
        // - a submission or a BETA board. Boards when nothing waits for this account in EITHER
        // queue - "awaiting" is the badge's own sense, so a submission waiting only for the other
        // approver does not keep the tab on a queue - else the screen on show.
        //
        // *** ONLY A QUEUE SCREEN GIVES WAY, AND NEVER OVER SOMETHING OPEN. *** A submission the
        // maintainer opened, which waits for the other approver, is still where it was on coming
        // back; and Boards and Account were chosen, so they stay. Switching would lose nothing - a
        // screen is hidden, never closed - but it would move the maintainer away from what they
        // were looking at.
        // ###########################################################################################
        public static MaintainerMode ScreenOnOpening(MaintainerMode shown, int attention, bool somethingOpen) =>
            attention <= 0 && !somethingOpen && shown is MaintainerMode.Review or MaintainerMode.Beta
                ? MaintainerMode.Boards
                : shown;

        // The same for the BETA list, whose entries are boards - and for the Boards list.
        public static string? EntryToOpen(IReadOnlyList<string> inListOrder, string? remembered)
        {
            ArgumentNullException.ThrowIfNull(inListOrder);

            if (remembered is not null && inListOrder.Contains(remembered, StringComparer.Ordinal))
                return remembered;

            return inListOrder.Count > 0 ? inListOrder[0] : null;
        }

        // The attention badge's text, or null to hide it: nothing needing you is no badge at all.
        public static string? AttentionBadge(int count) => TabBadge.Text(count);

        // ###########################################################################################
        // The Boards button's DISCREET badge: how many boards there are. Null - hidden - until the
        // list has been read, so a failed request is never shown as "0 boards".
        // ###########################################################################################
        public static string? CountBadge(int? count) =>
            count is int known && known >= 0 ? known.ToString(CultureInfo.InvariantCulture) : null;
    }
}
