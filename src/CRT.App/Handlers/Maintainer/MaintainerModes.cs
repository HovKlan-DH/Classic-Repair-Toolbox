using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // THE FOUR SCREENS BEHIND THE BUTTONS AT THE TOP LEFT (owner request, 2026-09-27):
    //
    //   Review   - the queue of contributions, as it always was.
    //   BETA     - systems in BETA waiting to be published to production (the "Production" window).
    //   Systems  - every system: its maintainers, contributors and submissions.
    //   Admin    - "Set maintainers" and "Unused files" (the administrator's two windows).
    //
    // Each shows its list on the left and what is chosen in it on the right, where three separate
    // windows used to open over the queue.
    //
    // *** THE BADGES COUNT SYSTEMS, NOT SUBMISSIONS ***, because that is what was asked for ("how
    // many systems need your attention") and what a maintainer works through: two submissions to one
    // board are reviewed together. "Needs your attention" is the SERVER's answer each time - the
    // queue's AwaitsYou, the production list's AwaitsYou - so a board waiting only for the other
    // approver is not counted. Pure, so the counting is tested.
    // ###########################################################################################
    public enum MaintainerMode
    {
        Review,
        Beta,
        Systems,
        Admin
    }

    public static class MaintainerModes
    {
        // The most a badge spells out; more reads as "99+" - a number that large is a backlog to
        // deal with, and the digits would push the other buttons out of the row.
        public const int BadgeCeiling = 99;

        // ###########################################################################################
        // Review: the systems with at least one submission waiting for THIS account. A submission the
        // server does not say about (an older server, null) counts as yours - the queue shows it as
        // yours too, undimmed.
        // ###########################################################################################
        public static int ReviewAttention(IEnumerable<ReviewQueueRow>? rows) =>
            rows is null
                ? 0
                : rows.Where(row => row.AwaitsYou != false)
                    .Select(row => row.SystemId ?? string.Empty)
                    .Distinct(StringComparer.Ordinal)
                    .Count();

        // BETA: the systems waiting for production that wait for THIS account - not the ones this
        // account has already approved, which wait for the other approver.
        public static int BetaAttention(IEnumerable<ProductionSystemRow>? rows) =>
            rows?.Count(row => row.AwaitsYou != false) ?? 0;

        // The attention badge's text, or null to hide it: nothing needing you is no badge at all.
        public static string? AttentionBadge(int count) =>
            count <= 0
                ? null
                : count > MaintainerModes.BadgeCeiling
                    ? $"{MaintainerModes.BadgeCeiling.ToString(CultureInfo.InvariantCulture)}+"
                    : count.ToString(CultureInfo.InvariantCulture);

        // ###########################################################################################
        // The Systems button's DISCREET badge: how many systems there are. Null - hidden - until the
        // list has been read, so a failed request is never shown as "0 systems".
        // ###########################################################################################
        public static string? CountBadge(int? count) =>
            count is int known && known >= 0 ? known.ToString(CultureInfo.InvariantCulture) : null;

        // The button's tooltip: what the screen is, then what its badge counts. `needingPlace` is the
        // Systems screen's: new systems this account can place in the drop-down lists (2026-09-27).
        public static string Tooltip(MaintainerMode mode, int? count, int needingPlace = 0)
        {
            string what = mode switch
            {
                MaintainerMode.Review => "Contributions waiting for review.",
                MaintainerMode.Beta => "Systems in BETA, waiting to be published to production or pushed back to the queue.",
                MaintainerMode.Systems => "Every system - who maintains it, who has contributed and how that went.",
                _ => "Who maintains which system, and files nothing uses."
            };

            string? counted = mode switch
            {
                MaintainerMode.Review => MaintainerModes.Systems(count, "with a contribution waiting for you", "with contributions waiting for you"),
                MaintainerMode.Beta => MaintainerModes.Systems(count, "waiting for you", "waiting for you"),
                MaintainerMode.Systems => count is int all ? MaintainerModes.Plural(all, "system", "systems") + " in all." : null,
                _ => null
            };

            if (mode == MaintainerMode.Systems && needingPlace > 0)
            {
                string waiting = needingPlace == 1
                    ? "1 new system needs a place in the drop-down lists."
                    : $"{needingPlace.ToString(CultureInfo.InvariantCulture)} new systems need a place in the drop-down lists.";

                counted = counted is null ? waiting : $"{waiting}\n{counted}";
            }

            return counted is null ? what : $"{what}\n{counted}";
        }

        // "Nothing is waiting for you." / "1 system ..." / "3 systems ...".
        private static string? Systems(int? count, string one, string many) =>
            count switch
            {
                null => null,
                <= 0 => "Nothing is waiting for you.",
                1 => $"1 system {one}.",
                _ => $"{count.Value.ToString(CultureInfo.InvariantCulture)} systems {many}."
            };

        private static string Plural(int count, string one, string many) =>
            count == 1 ? $"1 {one}" : $"{count.ToString(CultureInfo.InvariantCulture)} {many}";
    }
}
