namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A DRAFT HELD IN MEMORY SAYS ABOUT ITSELF (CRT 4.0 step 1, 2026-10-25 - the plan: "The
    // header says Saved, Saving, or "Could not save - <reason>" with Try again"). The Drafts tab's
    // table says it where "Save changes" was; the draft's header will say the same (step 1's last
    // slice). One place, so every editor of a draft words its state alike and the words are tested.
    // ###########################################################################################
    public static class DraftSaveWording
    {
        public const string Saved = "Saved";

        // Pending and Saving read the same: a change waiting its second is being saved.
        public const string Saving = "Saving...";

        public static string StateLine(DraftSaveState state, DraftWorkbookEditOutcome? failure) => state switch
        {
            DraftSaveState.Saved => DraftSaveWording.Saved,
            DraftSaveState.Failed => DraftSaveWording.CouldNotSave(failure),
            _ => DraftSaveWording.Saving,
        };

        public static string CouldNotSave(DraftWorkbookEditOutcome? failure) => "Could not save - " + failure switch
        {
            // Said on the draft's row too, table closed - so it names no table.
            DraftWorkbookEditOutcome.ChangedOnDisk => "the draft file was changed somewhere else.",
            DraftWorkbookEditOutcome.NoDraft => "this draft no longer exists.",
            _ => "the draft file could not be written. If another program has it open, close it there, then try again.",
        };

        // ###########################################################################################
        // Why another editor's save - the component editor, the label editor, the KiCad calibration,
        // "Schematic images" - could not be made; its change was taken back out (DraftEdits). The one
        // way the draft can be "changed outside" then: the table had a change waiting when the file
        // was written by another program, and the table asks which version to keep.
        // ###########################################################################################
        public static string EditNotSaved(DraftWorkbookEditOutcome? failure) => failure switch
        {
            DraftWorkbookEditOutcome.ChangedOnDisk =>
                "The draft was changed outside CRT while changes in the table view waited to be saved. " +
                "Choose there which version to keep, then save again.",
            DraftWorkbookEditOutcome.NoDraft => "This draft no longer exists.",
            _ => "The draft file could not be written. If another program has it open, close it there, then save again.",
        };

        // ###########################################################################################
        // The draft's file changed outside the table while a change of the table's waited to be
        // saved - so neither is thrown away unasked. The two buttons beside it say what each keeps.
        // ###########################################################################################
        public const string ChangedOutside =
            "The draft was changed outside this table - elsewhere in CRT, or by another program - while the table's " +
            "latest changes waited to be saved, so they were not saved over it. Use the draft as it was saved there " +
            "(Ctrl+Z then brings the table's version back), or keep the table's version and save it over the file.";

        public const string UseSavedVersion = "Use the saved version";

        public const string KeepTableVersion = "Keep the table's version";

        // The history's names for reading the file again (DraftSession.ReloadFromFile).
        public const string ReadAgainStep = "Changes made outside the table";

        public const string UsedSavedVersionStep = "The table's version, before the saved one was used";

        // What the status line says after each.
        public const string UpdatedFromOutside = "Updated - the draft was changed outside this table. Ctrl+Z takes that change back.";

        public const string UsedSavedVersion = "The table shows the draft as it was saved. Ctrl+Z brings back the table's version.";

        // ###########################################################################################
        // The draft header's Undo and Redo (step 1's last slice, 2026-10-25): what each would take
        // back or do again - one step of the draft's history, whichever editor made it.
        // ###########################################################################################
        public static string UndoTip(string? description) =>
            string.IsNullOrWhiteSpace(description) ? "Nothing to undo" : $"Undo: {description}";

        public static string RedoTip(string? description) =>
            string.IsNullOrWhiteSpace(description) ? "Nothing to redo" : $"Redo: {description}";

        public const string TryAgain = "Try again";
    }
}
