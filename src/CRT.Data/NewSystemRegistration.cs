namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE IDENTITY OF A SYSTEM THAT EXISTS ONLY AS A LOCAL DRAFT - one created through "Add a new
    // system" rather than one that arrived in the synced "Data" tree.
    //
    // *** IT LIVES IN THE DRAFT'S MARKER (.crt-draft.json), and that placement is deliberate. ***
    // Discarding a draft deletes the whole system FOLDER - "one folder is the whole draft", the
    // same model WorklogManager.DeleteWorkbook uses for a workbook - so a registration stored
    // inside it is retired atomically with the board and the files it belongs to. A separate
    // registry at the drafts root would be a second source of truth that the same delete could not
    // keep in step, leaving an entry pointing at a folder that no longer exists.
    //
    // Discovery walks the folder tree instead (DraftManager.EnumerateDraftOnlySystems), which is
    // self-validating: what is on disk IS the list.
    //
    // *** ExcelDataFile NAMES A FILE THAT DOES NOT EXIST IN Data/. *** It is the system's identity
    // key everywhere else in the app - the BoardDataReader cache key, the draft folder's own path,
    // the worklog board key - so a draft-only system needs one that has the same SHAPE as a
    // published system's even though nothing published stands behind it. See NewSystemIdentity for
    // how it is built and why a real-looking path is still the right choice.
    //
    // Moved out of BoardDraft.cs when the delta draft model was retired (Phase 6, 2026-09-23). The
    // record itself is unchanged; only the file that holds it and the file that stores it have
    // moved.
    // ###########################################################################################
    public sealed class NewSystemRegistration
    {
        public string HardwareName { get; init; } = string.Empty;

        public string BoardName { get; init; } = string.Empty;

        public string HardwareNotes { get; init; } = string.Empty;

        public string ExcelDataFile { get; init; } = string.Empty;

        public string CreatedUtc { get; init; } = string.Empty;

        // ###########################################################################################
        // The notes as a submission carries them (owner request, 2026-10-05 - they go to the server,
        // and from the maintainer's placement into the main Excel data file): trimmed, and cut to that
        // file's notes column. "Create system" stops typing there, but a draft made before it did,
        // or a marker edited by hand, could hold more - and CRT has nowhere to shorten them after
        // the system is created, so refusing them would leave the draft unsendable for good.
        // ###########################################################################################
        public string NotesForSubmission()
        {
            string notes = this.HardwareNotes?.Trim() ?? string.Empty;

            return notes.Length <= MasterListing.MaximumNotesLength
                ? notes
                : notes[..MasterListing.MaximumNotesLength].TrimEnd();
        }
    }
}
