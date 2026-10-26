using Handlers.DataHandling;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // EVERY SENTENCE THE DRIFT WARNING SAYS, in one place (code review, 2026-10-13) - the Drafts
    // row's line, the main window's notice and the "What changed in the published data" window.
    //
    // *** WHAT CHANGED, AND WHY. *** All three used to say "your edits are still applied on top of
    // it". That had stopped being true when drafts became complete copies (Phase 6): a draft started
    // before somebody else's work was published does not hold that work, and sending it would show
    // the maintainer every row published since as a removal by this contributor. A draft seeded
    // since 2026-10-13 keeps a base and CAN be brought up to date (DraftRebaser) - so the words say
    // so and offer it; an older one cannot, and the words say THAT plainly instead.
    //
    // "The published data" is the one name for the board as CRT downloads it (code review: it was
    // "official data", "online source" and "published data" in three places).
    // ###########################################################################################
    public static class DraftDriftWording
    {
        public static bool IsWarning(DraftDriftState state) =>
            state is DraftDriftState.OfficialIsNewer or DraftDriftState.Changed;

        // "has been updated" only when the published revision is known to be later.
        private static string What(DraftDriftState state) =>
            state == DraftDriftState.OfficialIsNewer ? "has been updated" : "has changed";

        // The window's opening paragraph.
        public static string Summary(DraftDriftState state, bool canBringUpToDate) => state switch
        {
            DraftDriftState.OfficialIsNewer or DraftDriftState.Changed when canBringUpToDate =>
                $"The published data for this board {What(state)} since you started your draft, and your draft does not " +
                "have that yet. Bring your draft up to date to add it - every change of yours is kept.",
            DraftDriftState.OfficialIsNewer or DraftDriftState.Changed =>
                $"The published data for this board {What(state)} since you started your draft, and your draft does not " +
                "have that update. This draft was started before CRT could bring a draft up to date, so if you send it as " +
                "it is, the maintainer sees the update's changes as changes you made. If you changed only a little, the " +
                "safe way is to discard this draft and make your changes again.",
            DraftDriftState.InSync =>
                "The published data for this board has not changed since you started your draft.",
            _ =>
                "CRT does not know which revision this draft was started from, so it cannot tell whether the published data has changed.",
        };

        // One line for the Drafts row and the main window's notice.
        public static string Line(DraftDriftState state, bool canBringUpToDate) => canBringUpToDate
            ? $"The published data for this board {What(state)} since you started your draft. Bring your draft up to date before you send it."
            : $"The published data for this board {What(state)} since you started your draft, and your draft does not have the update.";

        // What the lists in the window compare.
        public static string Compared(bool comparedWithBase) => comparedWithBase
            ? "Your changes are measured against the published data as it was when you started your draft, so your work " +
              "and the update are listed apart."
            : "This draft was started before CRT kept a copy of the data it was made from, so it can only be compared " +
              "with the published data as it is now - and that includes the update's changes, turned round.";

        public static string OwnHeader(int count, bool comparedWithBase) => comparedWithBase
            ? (count == 1 ? "1 change of yours" : $"{count} changes of yours")
            : (count == 1 ? "1 difference from the published data" : $"{count} differences from the published data");

        public static string UpdateHeader(int count) =>
            count == 1 ? "1 change in the published update" : $"{count} changes in the published update";

        public static string ConflictHeader(int count) =>
            count == 1 ? "1 row changed by both you and the update" : $"{count} rows changed by both you and the update";

        // A row of the contributor's, said as what they did (with a base) or as the plain difference.
        public static string OwnRow(BoardRowChange row, bool comparedWithBase) => (row.Kind, comparedWithBase) switch
        {
            (BoardRowChangeKind.Added, true) => "You added this.",
            (BoardRowChangeKind.Deleted, true) => "You removed this - your submission would remove it from the published data.",
            (BoardRowChangeKind.Added, false) => "In your draft, not in the published data.",
            (BoardRowChangeKind.Deleted, false) => "In the published data, not in your draft - your submission would remove it.",
            (_, true) => row.ChangedFields.Count > 0 ? "You changed: " + string.Join(", ", row.ChangedFields) : "You changed this.",
            _ => row.ChangedFields.Count > 0 ? "Differs from the published data: " + string.Join(", ", row.ChangedFields) : "Differs from the published data.",
        };

        public static string UpdateRow(BoardRowChange row) => row.Kind switch
        {
            BoardRowChangeKind.Added => "Added in the update.",
            BoardRowChangeKind.Deleted => "Removed in the update.",
            _ => row.ChangedFields.Count > 0 ? "Changed in the update: " + string.Join(", ", row.ChangedFields) : "Changed in the update.",
        };

        public const string BringUpToDateButton = "Bring my draft up to date";

        public const string BringUpToDateTip = "Adds the published update to your draft. Every change of yours is kept.";

        public const string SeenButton = "I have looked at this";

        public const string SeenTip = "Stop warning about this update. Your draft is not changed - it still does not have the update.";

        public const string SeenDone = "Noted. Your draft is unchanged, and it still does not have the update.";

        // What "Bring my draft up to date" did.
        public static string Outcome(DraftRebaseReport report) => report.Outcome switch
        {
            DraftRebaseOutcome.BroughtUpToDate when report.Conflicts.Count > 0 =>
                $"Your draft now has the published update. {Kept(report.ChangesKept)} " +
                $"{report.Conflicts.Count} of them {(report.Conflicts.Count == 1 ? "was" : "were")} also changed by the update - " +
                "your version was kept. Look at the list above before you send.",
            DraftRebaseOutcome.BroughtUpToDate =>
                $"Your draft now has the published update. {Kept(report.ChangesKept)}",
            DraftRebaseOutcome.AlreadyUpToDate =>
                "Your draft already has everything that is published.",
            DraftRebaseOutcome.NotNoted =>
                "Your draft now has the published update, but CRT could not note that it does, so the warning stays. " +
                "Press \"" + DraftDriftWording.BringUpToDateButton + "\" again to finish - nothing of yours is lost.",
            DraftRebaseOutcome.ChangedOnDisk =>
                "Your draft changed while this ran (saved in another window, perhaps). Nothing was changed - try again.",
            DraftRebaseOutcome.WriteFailed =>
                "Your draft could not be written - is one of its files open in another program? Close it and try again. Nothing was changed.",
            DraftRebaseOutcome.PublishedUnreadable =>
                "The published data for this board could not be read. Nothing was changed.",
            DraftRebaseOutcome.DraftUnreadable =>
                "Your draft could not be read. Nothing was changed.",
            _ =>
                "This draft cannot be brought up to date. Nothing was changed.",
        };

        public static bool IsSuccess(DraftRebaseReport report) =>
            report.Outcome is DraftRebaseOutcome.BroughtUpToDate or DraftRebaseOutcome.AlreadyUpToDate;

        private static string Kept(int count) => count switch
        {
            0 => "You had no changes of your own.",
            1 => "Your 1 change was kept.",
            _ => $"All {count} changes of yours were kept.",
        };
    }
}
