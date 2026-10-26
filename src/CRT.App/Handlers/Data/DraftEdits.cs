using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE EDITOR'S CHANGE TO A DRAFT, THROUGH THE DRAFT HELD IN MEMORY (CRT 4.0 step 1, 2026-10-25 -
    // the plan: "Every editor works on the same draft, held once in memory. An edit in one shows up in
    // all the others at once. Undo and redo work across every editor").
    //
    // The component editor, the label editor, the KiCad calibration and "Schematic images" used to
    // read the draft's file, change it and write it back each on its own - and had to hold each other
    // back from writing under the Drafts tab's table. Now each hands its change to the draft's one
    // DraftSession (DraftManager.Sessions):
    //
    //   - APPLIED IN MEMORY, so the table, if open, shows it at once (it hands over what it holds
    //     first - DraftSession.ChangeComing) and its Ctrl+Z takes it back;
    //   - SAVED STRAIGHT AWAY, off the UI thread - the editor's own Save button, as before;
    //   - TAKEN BACK OUT when that save cannot be made (DraftSession.TakeBack), so an editor's save
    //     either happens or changes nothing, and it says "could not be saved" as it always did.
    //
    // Something written to the file by hand meanwhile is read first (ReadFileIfChanged), so the change
    // lands on the draft as it is. UI THREAD ONLY: applying raises the session's events to the table.
    // ###########################################################################################
    internal static class DraftEdits
    {
        public static Task<DraftWorkbookEditOutcome> ApplyAndSaveAsync(
            string excelDataFile,
            string description,
            Func<BoardContent, BoardContent?> edit,
            IReadOnlyDictionary<string, string?>? follow = null,
            bool takeInNewFiles = false)
        {
            ArgumentNullException.ThrowIfNull(edit);

            return DraftEdits.ApplyWithFollowAndSaveAsync(excelDataFile, description, content => (edit(content), follow), takeInNewFiles);
        }

        // The same for a change of the board's rows and highlights alone - what every editor but the
        // calibration changes. A null board, or the board it was handed, is nothing to do - never an
        // empty step in the history.
        //
        // `takeInNewFiles`: the editor copied files into the draft folder for this change - a change
        // that changes no row is then still written, so they are taken in now (a picture replaced
        // under the name its row already cites; code review, 2026-10-26).
        public static Task<DraftWorkbookEditOutcome> ApplyAndSaveAsync(
            string excelDataFile,
            string description,
            Func<BoardData, BoardData?> transform,
            IReadOnlyDictionary<string, string?>? follow = null,
            bool takeInNewFiles = false)
        {
            ArgumentNullException.ThrowIfNull(transform);

            return DraftEdits.ApplyAndSaveAsync(
                excelDataFile,
                description,
                content => transform(content.Board) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null,
                follow,
                takeInNewFiles);
        }

        // ###########################################################################################
        // "UNDO CHANGES TO <component>" (step 1's last slice, 2026-10-25): every row of the component
        // back to `published`'s - gone, for a component the published board does not have - as ONE
        // step of the draft's history, so even this is one Undo away (ComponentRevert). "Review
        // changes" asks it; the component pane does the same on its own session.
        //
        // *** AND ITS FILES REPLACED UNDER THEIR OWN NAMES (code review, 2026-10-26). *** The rows name
        // them the same, so the draft's list kept the contributor's bytes and the next submission sent
        // them: those go back to the published board's (`publishedBoardFile`, "" for none) once the
        // rows are saved - a write of the file list, which the draft's Undo does not take back.
        // ###########################################################################################
        public static async Task<DraftWorkbookEditOutcome> UndoComponentChangesAsync(string excelDataFile, string label, BoardData? published, string publishedBoardFile)
        {
            string wanted = label.Trim();
            string description = DraftStepWording.ComponentReverted(wanted);
            string draftsRoot = DraftManager.DraftsRoot;

            // Asked on the draft as it is BEFORE its rows go back: a relabelled component's files are
            // the ones its old label cites.
            DraftSession? session = DraftManager.Sessions.Find(draftsRoot, excelDataFile)
                                    ?? await Task.Run(() => DraftManager.Sessions.Get(draftsRoot, excelDataFile));

            List<string> replaced = session is null
                ? []
                : await Task.Run(() => ComponentRevert.ReplacedFiles(session.Content.Board, published, wanted, DraftFiles.ListedIn(publishedBoardFile), session.ListedFiles));

            DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                excelDataFile,
                description,
                (BoardData board) => ComponentRevert.RowsToPublished(board, published, wanted));

            return outcome != DraftWorkbookEditOutcome.Saved || replaced.Count == 0 || session is null
                ? outcome
                : await DraftEdits.PutBackPublishedFilesAsync(session, replaced, publishedBoardFile, description);
        }

        // ###########################################################################################
        // The files `names` of the draft's list back to the published board's (`publishedBoardFile`) -
        // the bytes the draft stored for them go (DraftWorkbookStore.RemoveFiles). A write of the file
        // list, so it runs under the session's lock, what waits saved first and the file read back
        // (WriteFileDirectly). Saved when done; the failure otherwise. UI thread.
        // ###########################################################################################
        public static async Task<DraftWorkbookEditOutcome> PutBackPublishedFilesAsync(
            DraftSession session,
            IReadOnlyList<string> names,
            string publishedBoardFile,
            string description)
        {
            ArgumentNullException.ThrowIfNull(session);
            ArgumentNullException.ThrowIfNull(names);

            session.CollectHeldChanges(description);

            try
            {
                bool done = await Task.Run(() => session.WriteFileDirectly(
                    () => DraftWorkbookStore.RemoveFiles(session.DraftsRoot, session.ExcelDataFile, names, publishedBoardFile),
                    succeeded => succeeded,
                    description,
                    clearHistory: false));

                return done ? DraftWorkbookEditOutcome.Saved : DraftWorkbookEditOutcome.WriteFailed;
            }
            catch (DraftNotSavedException ex)
            {
                return ex.Failure;
            }
        }

        // ###########################################################################################
        // The same for a change that works out its schematic renames on the draft as it is (a rename
        // decided by SchematicRename.Plan - DraftSession.ApplyWithFollow). Saved off the UI thread.
        // ###########################################################################################
        public static async Task<DraftWorkbookEditOutcome> ApplyWithFollowAndSaveAsync(
            string excelDataFile,
            string description,
            Func<BoardContent, (BoardContent? Content, IReadOnlyDictionary<string, string?>? Follow)> edit,
            bool takeInNewFiles = false)
        {
            if (await DraftEdits.ApplyAsync(excelDataFile, description, edit) is not (DraftSession session, object origin))
            {
                // No step to take back, so a write that fails changes nothing either.
                return takeInNewFiles && DraftManager.Sessions.Find(DraftManager.DraftsRoot, excelDataFile) is DraftSession unchanged
                    ? await unchanged.SaveAsync(evenUnchanged: true)
                    : DraftEdits.NothingToSave(excelDataFile);
            }

            DraftWorkbookEditOutcome outcome = await session.SaveAsync();

            return DraftEdits.KeptOrTakenBack(session, origin, outcome);
        }

        // ###########################################################################################
        // The same, saved on the calling thread - only where nothing may wait: a window closed by a
        // shutdown of CRT or the computer ("Schematic images" with a name still typed).
        // ###########################################################################################
        public static DraftWorkbookEditOutcome ApplyWithFollowAndSaveNow(
            string excelDataFile,
            string description,
            Func<BoardContent, (BoardContent? Content, IReadOnlyDictionary<string, string?>? Follow)> edit)
        {
            if (DraftEdits.Apply(excelDataFile, description, edit) is not (DraftSession session, object origin))
            {
                return DraftEdits.NothingToSave(excelDataFile);
            }

            DraftWorkbookEditOutcome outcome;

            try
            {
                outcome = session.Save();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Saving the draft for [{excelDataFile}] failed unexpectedly - [{ex}]");
                outcome = DraftWorkbookEditOutcome.WriteFailed;
            }

            return DraftEdits.KeptOrTakenBack(session, origin, outcome);
        }

        // ###########################################################################################
        // The change applied to the draft's session, with its own mark - or null when there is no
        // draft, or the change changed nothing.
        //
        // *** THE HEAVY PART OFF THE UI THREAD (code review, 2026-10-25). *** Opening the session reads
        // and hashes the board file, asking whether it changed hashes it again, and the edit and what
        // follows it (DraftEditRules) walk the whole board - all of which ran on the UI thread before
        // the editor's "please wait" could paint. Now the session is opened and the file asked about on
        // the pool, and the edit is worked out there on the draft as it is (DraftSession.Prepare); only
        // handing over what the table holds, and applying the result, stay on the UI thread. A draft
        // that moved in between has the edit worked out again, there.
        // ###########################################################################################
        private static async Task<(DraftSession Session, object Origin)?> ApplyAsync(
            string excelDataFile,
            string description,
            Func<BoardContent, (BoardContent? Content, IReadOnlyDictionary<string, string?>? Follow)> edit)
        {
            ArgumentNullException.ThrowIfNull(edit);

            if (await DraftEdits.OpenForChangeAsync(excelDataFile) is not DraftSession session)
            {
                return null;
            }

            // The change's own mark, so a save that cannot be made takes back this change and no other.
            object origin = new DraftEditOrigin(description);

            BoardContent before = session.Content;

            if (await Task.Run(() => DraftSession.Prepare(before, edit)) is not DraftPreparedEdit prepared)
                return null;

            bool applied = session.ApplyPrepared(description, prepared, origin) ||
                           session.ApplyWithFollow(description, edit, origin);

            return applied ? (session, origin) : null;
        }

        // ###########################################################################################
        // The draft's session, opened for a change as every editor's change opens it: on the pool when
        // it is not open yet, what its editors hold handed over first (so the file is asked about with
        // it in the session), and a file written meanwhile read - the change then lands on the draft as
        // it is. Null without a draft. UI thread.
        // ###########################################################################################
        public static async Task<DraftSession?> OpenForChangeAsync(string excelDataFile)
        {
            string draftsRoot = DraftManager.DraftsRoot;

            if ((DraftManager.Sessions.Find(draftsRoot, excelDataFile)
                 ?? await Task.Run(() => DraftManager.Sessions.Get(draftsRoot, excelDataFile))) is not DraftSession session)
            {
                return null;
            }

            session.CollectHeldChanges(DraftSaveWording.ReadAgainStep);

            if (await Task.Run(session.FileChangedOutside))
                session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep);

            return session;
        }

        // ###########################################################################################
        // `files` - each a file and the name the change gives it - copied into the draft's folder, then
        // `transform` - the rows citing them - applied as ONE step and saved, as every editor's change
        // (ApplyAndSaveAsync: taken back out when the save cannot be made). The copies the write took
        // in are kept; when the save is not made, or the write did not take a copy in, the copy is
        // taken back out with what it replaced (DraftPickedFiles.KeepTakenIn; code review, 2026-10-26 -
        // left lying there, the next write by any editor took it in unasked). The save's outcome; null
        // when nothing could be copied. UI thread, as every change.
        // ###########################################################################################
        public static async Task<DraftWorkbookEditOutcome?> CopyInAndApplyAsync(
            string excelDataFile,
            IReadOnlyList<(string Source, string Name)> files,
            string description,
            Func<BoardData, BoardData?> transform)
        {
            ArgumentNullException.ThrowIfNull(files);
            ArgumentNullException.ThrowIfNull(transform);

            string folder = DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, excelDataFile);
            DraftPickedFiles picked = await Task.Run(() => DraftPickedFiles.CopyInto(folder, files));

            if (!picked.AnyLying)
            {
                return null;
            }

            DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(excelDataFile, description, transform, takeInNewFiles: true);

            if (outcome == DraftWorkbookEditOutcome.Saved)
            {
                picked.KeepTakenIn();
            }
            else
            {
                picked.TakeBack();
            }

            return outcome;
        }

        // The same, all on the calling thread - for ApplyWithFollowAndSaveNow, which nothing may wait for.
        private static (DraftSession Session, object Origin)? Apply(
            string excelDataFile,
            string description,
            Func<BoardContent, (BoardContent? Content, IReadOnlyDictionary<string, string?>? Follow)> edit)
        {
            ArgumentNullException.ThrowIfNull(edit);

            if (DraftManager.Sessions.Get(DraftManager.DraftsRoot, excelDataFile) is not DraftSession session)
            {
                return null;
            }

            session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep);

            // The change's own mark, so a save that cannot be made takes back this change and no other.
            object origin = new DraftEditOrigin(description);

            return session.ApplyWithFollow(description, edit, origin) ? (session, origin) : null;
        }

        // Nothing applied: no draft at all, or a change that changed nothing - which is saved already.
        private static DraftWorkbookEditOutcome NothingToSave(string excelDataFile) =>
            DraftManager.Sessions.Find(DraftManager.DraftsRoot, excelDataFile) is null
                ? DraftWorkbookEditOutcome.NoDraft
                : DraftWorkbookEditOutcome.Saved;

        // A change whose save could not be made, taken back. When it cannot be - its step no longer in
        // the history at all (MaxSteps of changes since, or the history cleared) - the change stays in
        // the draft to be saved with the next change, and the log says so: the editor still says it
        // could not be saved, which is true of this save.
        private static DraftWorkbookEditOutcome KeptOrTakenBack(DraftSession session, object origin, DraftWorkbookEditOutcome outcome)
        {
            if (outcome != DraftWorkbookEditOutcome.Saved && !session.TakeBack(origin))
            {
                Logger.Warning($"A change to the draft of [{session.ExcelDataFile}] could not be saved ({outcome}) nor taken back - its step is no longer in the history; it stays, for the next save");
            }

            return outcome;
        }

        // ###########################################################################################
        // A WRITE OF THE DRAFT'S FILE THAT IS MORE THAN ITS CONTENT - "Bring my draft up to date",
        // "Remove" of an imported KiCad file - run on the pool under its session's lock when one is
        // open (DraftSession.WriteFileDirectly: saved first, the file read back after), else on its
        // own. Called on the UI thread, where an editor holding a change hands it over first
        // (CollectHeldChanges); the window asking holds CRT's main window while it runs. A draft whose
        // changes cannot be saved first throws DraftNotSavedException, and `write` does not run.
        // ###########################################################################################
        public static Task<T> WriteFileDirectlyAsync<T>(string excelDataFile, Func<T> write, Func<T, bool> succeeded, string description, bool clearHistory)
        {
            ArgumentNullException.ThrowIfNull(write);

            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, excelDataFile) is not DraftSession session)
            {
                return Task.Run(write);
            }

            session.CollectHeldChanges(description);

            return Task.Run(() => session.WriteFileDirectly(write, succeeded, description, clearHistory));
        }

        // ###########################################################################################
        // The board file written with nothing changed in it, so the new files lying in the draft folder
        // are taken in and listed - after a KiCad project was copied in. A draft with no session open
        // is written as before (DraftWorkbookStore.Edit). False when it could not be written.
        // ###########################################################################################
        public static async Task<bool> TakeInNewFilesAsync(string excelDataFile)
        {
            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, excelDataFile) is DraftSession session)
            {
                return await session.SaveAsync(evenUnchanged: true) == DraftWorkbookEditOutcome.Saved;
            }

            string draftsRoot = DraftManager.DraftsRoot;
            return await Task.Run(() => DraftWorkbookStore.Edit(draftsRoot, excelDataFile, board => board));
        }
    }

    // The mark one editor's change carries in the draft's history - see DraftEdits.ApplyAndSaveAsync.
    internal sealed record DraftEditOrigin(string Description);
}
