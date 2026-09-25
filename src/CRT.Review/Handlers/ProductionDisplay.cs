using System;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace CRT.Review.Handlers
{
    // ###########################################################################################
    // How the "Production" window reads (maintainer request, 2026-09-25: BETA first, then
    // production "only after he has checked that the data looks correct in BETA").
    //
    // Pure, so the words are tested - the same rule ReviewQueueDisplay and
    // ReviewerAssignmentDisplay follow. One rule here is a decision rather than formatting: the
    // "Publish to production" button follows the SERVER's canPublish AND the reviewer's own tick in
    // "I have checked this in CRT with the BETA data" - never either alone (CanPress).
    // ###########################################################################################
    public static class ProductionDisplay
    {
        // "Commodore / C64 / 250407  -  BETA 2026-September-25, production 2026-May-14" - the
        // heading over the plan, where it has the width to wrap.
        public static string SystemLine(ProductionSystemRow system) =>
            $"{ProductionDisplay.SystemName(system)}  -  {ProductionDisplay.SystemStatus(system)}";

        // ###########################################################################################
        // The same, on TWO lines, for the list on the left. One line was cut off at the list's
        // width in the first render - "...never p" - which hid the one fact the row is there to
        // give: how far production is behind.
        // ###########################################################################################
        public static string ListEntry(ProductionSystemRow system) =>
            $"{ProductionDisplay.SystemName(system)}\n{ProductionDisplay.SystemStatus(system)}";

        public static string SystemName(ProductionSystemRow system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string name = string.Join(
                " / ",
                new[] { system.Manufacturer, system.Hardware, system.Board }
                    .Where(part => !string.IsNullOrWhiteSpace(part)));

            return name.Length > 0
                ? name
                : string.IsNullOrWhiteSpace(system.SystemId) ? "(unknown system)" : system.SystemId;
        }

        // "BETA 2026-September-25, production 2026-May-14"
        public static string SystemStatus(ProductionSystemRow system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string beta = string.IsNullOrWhiteSpace(system.BetaRevision) ? "BETA ahead" : $"BETA {system.BetaRevision}";

            string production = string.IsNullOrWhiteSpace(system.ProductionRevision)
                ? "never published to production"
                : $"production {system.ProductionRevision}";

            return $"{beta}, {production}";
        }

        // ###########################################################################################
        // "2 files to add, 1 to replace, 1,180 already the same". Nothing to copy is said as such,
        // because it is a real outcome (production was brought level by hand) and publishing then
        // only records it.
        // ###########################################################################################
        public static string PlanSummary(ProductionPlanView plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            int added = plan.Files.Count(file => file.Change == PromotionChange.Added);
            int replaced = plan.Files.Count(file => file.Change == PromotionChange.Replaced);
            int removed = plan.Removals?.Files.Count ?? 0;

            string unchanged = plan.UnchangedCount.ToString("N0", CultureInfo.InvariantCulture);

            // Removals are said only when there are any: most promotions remove nothing, and a
            // standing "0 to remove" would be read past on the one that does.
            string removing = removed == 0 ? string.Empty : $", {ProductionDisplay.Count(removed, "file", "to remove")}";

            if (added == 0 && replaced == 0)
                return $"Production already has every file ({unchanged}){removing}. Publishing only records it.";

            return $"{ProductionDisplay.Count(added, "file", "to add")}, " +
                $"{ProductionDisplay.Count(replaced, "file", "to replace")}{removing}, {unchanged} already the same";
        }

        // "added     Commodore/C64/250407/Images/a.png" - with "(shared)" on a file that reaches
        // every board citing it, because that is the line a reviewer most needs to notice.
        public static string FileLine(PromotionFile file)
        {
            ArgumentNullException.ThrowIfNull(file);

            string change = file.Change == PromotionChange.Added ? "added" : "replaced";

            return file.IsShared ? $"{change}  {file.Path}  (shared)" : $"{change}  {file.Path}";
        }

        // ###########################################################################################
        // May the button be pressed? The server must say yes AND the reviewer must have ticked the
        // box. The tick is the human half of "only after he has checked it in BETA"; the server's
        // half is the content hash the request carries back.
        // ###########################################################################################
        public static bool CanPress(ProductionPlanView? plan, bool checkedInBeta) =>
            plan is not null && plan.CanPublish && checkedInBeta;

        private static string Count(int count, string noun, string what) =>
            count == 1
                ? $"1 {noun} {what}"
                : $"{count.ToString("N0", CultureInfo.InvariantCulture)} {noun}s {what}";
    }
}
