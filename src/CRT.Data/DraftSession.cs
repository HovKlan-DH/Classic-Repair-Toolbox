using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A DRAFT HELD ONCE IN MEMORY, SHARED BY EVERY EDITOR, AND SAVED BY ITSELF - the foundation of
    // CRT 4.0's step 1, "Nothing is ever lost" (the plan of 2026-10-24: "Every change is saved within
    // a second ... Every editor works on the same draft, held once in memory. An edit in one shows up
    // in all the others at once ... Undo and redo work across every editor for the whole session").
    //
    // Until now every editor read the draft's board file, changed it and wrote it back
    // (DraftWorkbookStore), each with its own copy on screen and its own Save - so an edit in the table
    // was not in the component editor until something re-read the file, and two editors had to block
    // each other's saves. Here an edit is a transform of the draft as it is IN MEMORY (Apply), seen at
    // once by everybody listening (Changed), and written to the file a moment later by itself (Save,
    // which the host calls once IsSaveDue says so).
    //
    //   - ONE HISTORY FOR EVERY EDITOR. Each change is a step holding the draft as it was before -
    //     the whole content, which a transform never mutates (every draft writer returns a new
    //     board) - so undo puts back exactly what was there, whichever editor made the change, and
    //     says what it was (UndoDescription). MaxSteps bounds it.
    //   - WHAT FOLLOWS AN EDIT, FOLLOWS IT HERE TOO: highlights and calibrations go with a renamed
    //     schematic or a relabelled component - DraftEditRules, the rules every file write applies.
    //   - SAVED MEANS THE FILE HOLDS WHAT IS SHOWN: compared by the content itself, so undoing back
    //     to the content last written is Saved again, with nothing to write.
    //   - A FILE CHANGED OUTSIDE IS NEVER WRITTEN OVER WITHOUT BEING ASKED. The session keeps the
    //     fingerprint of the file as it read or last wrote it, and a save finding another refuses
    //     (Failed, ChangedOnDisk) - then the host offers the file's version (ReloadFromFile, itself an
    //     undoable step) or the one on screen (Save with overwriteOutsideChange).
    //   - A FAILED SAVE IS NOT RETRIED IN A LOOP. It is due again after the next change, or when the
    //     host retries ("Try again").
    //
    //   - BEFORE ANY CHANGE, WHOEVER HOLDS ONE NOT HANDED OVER YET HANDS IT OVER (ChangeComing): the
    //     Drafts tab's table pushes what it holds, so another editor's change lands on top of it and
    //     the table, built again afterwards, loses nothing.
    //   - AN EDITOR'S EXPLICIT SAVE EITHER HAPPENS OR CHANGES NOTHING: a change whose save could not be
    //     made is taken back out again (TakeBack) - the component editor and the label editor say
    //     "could not be saved", as they always did.
    //   - A WRITE THAT IS MORE THAN THE CONTENT (the board's file list, a draft brought up to date with
    //     its base and marker) runs under the session's own lock, everything saved first and the file
    //     read back after (WriteFileDirectly).
    //
    // THREADS: Apply, Undo, Redo, TakeBack and ReloadFromFile belong to the UI thread, and raise
    // ChangeComing and Changed there. Save and WriteFileDirectly run anywhere - the host writes off the
    // UI thread - and raise SaveStateChanged on the thread they run on. One write at a time.
    //
    // Avalonia-free, in CRT.Data like the rest of the draft model, so the rules are tested here.
    // ###########################################################################################
    public sealed class DraftSession
    {
        public const int MaxSteps = 200;

        // How long after the last change a save waits - "every change is saved within a second".
        public static readonly TimeSpan SaveDelay = TimeSpan.FromSeconds(1);

        private readonly object thisGate = new();
        private readonly object thisWriting = new();
        private readonly List<Step> thisUndo = new();
        private readonly List<Step> thisRedo = new();
        private readonly Func<DateTimeOffset> thisClock;

        // The files the board file listed while held here - so an undo bringing back a row citing a
        // file a save in between dropped finds it listed again, its bytes kept (DraftFileMemory).
        private readonly DraftFileMemory thisFiles = new();

        private BoardContent thisContent;
        private BoardContent thisSavedContent;
        private string thisFingerprint;

        // The file's stamp (length and write time) when it was last seen to hold thisFingerprint,
        // settled - see FileMoved. Null when not known; every new fingerprint clears it.
        private string? thisSettledStamp;
        private int thisFingerprintsTaken;
        private DateTimeOffset thisLastChange;
        private bool thisSaving;
        private DraftWorkbookEditOutcome? thisFailure;
        private BoardContent? thisFailedContent;

        private DraftSession(string draftsRoot, string excelDataFile, BoardContent content, string fingerprint, Func<DateTimeOffset> clock)
        {
            this.DraftsRoot = draftsRoot;
            this.ExcelDataFile = excelDataFile;
            this.thisContent = content;
            this.thisSavedContent = content;
            this.thisFingerprint = fingerprint;
            this.thisClock = clock;
            this.thisLastChange = clock();
        }

        // ###########################################################################################
        // The draft of `excelDataFile` under `draftsRoot`, its board file read once - or null when
        // there is no draft, or its board file cannot be read in full. `clock` is for tests.
        // ###########################################################################################
        public static DraftSession? Open(string draftsRoot, string excelDataFile, Func<DateTimeOffset>? clock = null)
        {
            string boardFile = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            if (boardFile.Length == 0 || !File.Exists(boardFile))
            {
                return null;
            }

            // The fingerprint FIRST: a write landing between the two reads then makes the next save
            // refuse rather than write over it.
            string fingerprint = DraftWorkbookStore.Fingerprint(draftsRoot, excelDataFile);

            if (BoardFiles.ReadDocument(boardFile) is not BoardDocument document || fingerprint.Length == 0)
            {
                return null;
            }

            var session = new DraftSession(draftsRoot, excelDataFile, BoardDocumentMapping.ToContent(document), fingerprint, clock ?? (() => DateTimeOffset.UtcNow));
            session.thisFiles.Note(document.Files);

            return session;
        }

        // The board file's content, its file list noted - null when it cannot be read.
        private BoardContent? ReadBoardFile()
        {
            if (BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, this.ExcelDataFile)) is not BoardDocument document)
            {
                return null;
            }

            this.thisFiles.Note(document.Files);

            return BoardDocumentMapping.ToContent(document);
        }

        public string DraftsRoot { get; }

        public string ExcelDataFile { get; }

        // The draft as every editor should show it now. Never mutate it - edit through Apply.
        public BoardContent Content
        {
            get
            {
                lock (this.thisGate)
                    return this.thisContent;
            }
        }

        // ###########################################################################################
        // The files the board file lists, by name (ignoring capitals), as this session last read or
        // wrote it - what an editor compares with the published board's (ComponentRevert.ReplacedFiles)
        // without reading a board file of a few megabytes again on the UI thread. The list follows the
        // rows at each save (DraftFiles), so between a change and its save it is the list before.
        // ###########################################################################################
        public IReadOnlyDictionary<string, BoardDocument.StoredFileRow> ListedFiles => this.thisFiles.Listed;

        // Raised on the UI thread after the content changed - an edit, an undo, a redo, the file read
        // again - saying what and by whom (`Origin`, the editor that made an edit, so it can skip
        // re-reading its own change).
        public event EventHandler<DraftSessionChange>? Changed;

        // Raised when SaveState may have changed - on the thread the change happened on.
        public event EventHandler? SaveStateChanged;

        // ###########################################################################################
        // Raised just BEFORE the content changes - an edit, an undo or redo, a take-back, the file read
        // again - saying what is coming and from whom. An editor holding a change it has not handed
        // over yet (the Drafts tab's table) hands it over now, so the coming change lands on top of it.
        // An editor ignores its own.
        // ###########################################################################################
        public event EventHandler<DraftSessionChange>? ChangeComing;

        // ###########################################################################################
        // Applies an edit to the draft in memory: `edit` gets the content as it is and returns the
        // content as it should be (a new object - never its input changed), or null, or the same
        // object, for nothing to do. What follows the edit follows it (DraftEditRules; `follow` is an
        // edit's own schematic renames when it knows them). One undo step, described for the history
        // ("Undo: <description>"); the redo steps go. True when the draft changed.
        // ###########################################################################################
        public bool Apply(
            string description,
            Func<BoardContent, BoardContent?> edit,
            object? origin = null,
            IReadOnlyDictionary<string, string?>? follow = null)
        {
            ArgumentNullException.ThrowIfNull(edit);

            return this.ApplyWithFollow(description, content => (edit(content), follow), origin);
        }

        // ###########################################################################################
        // The same, for an edit that knows its schematic renames only once it sees the draft - a
        // rename decided on the draft as it is (SchematicRename.Plan): `edit` returns the content and
        // the renames together.
        // ###########################################################################################
        public bool ApplyWithFollow(
            string description,
            Func<BoardContent, (BoardContent? Content, IReadOnlyDictionary<string, string?>? Follow)> edit,
            object? origin = null)
        {
            ArgumentNullException.ThrowIfNull(edit);

            this.ChangeComing?.Invoke(this, new DraftSessionChange(DraftSessionChangeKind.Edit, description ?? string.Empty, origin));

            return DraftSession.Prepare(this.Content, edit) is DraftPreparedEdit prepared &&
                   this.Commit(description, prepared, origin);
        }

        // ###########################################################################################
        // AN EDIT WORKED OUT OFF THE UI THREAD (code review, 2026-10-25: a large board's edit, and what
        // follows it, ran on the UI thread before the "please wait" could paint). `edit` is run on
        // `before` and what follows it is followed (DraftEditRules) - pure, so any thread. Null when
        // it changes nothing. ApplyPrepared then applies it on the UI thread.
        // ###########################################################################################
        public static DraftPreparedEdit? Prepare(
            BoardContent before,
            Func<BoardContent, (BoardContent? Content, IReadOnlyDictionary<string, string?>? Follow)> edit)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(edit);

            (BoardContent? edited, IReadOnlyDictionary<string, string?>? follow) = edit(before);

            if (edited is null || ReferenceEquals(edited, before))
            {
                return null;
            }

            (BoardData board, List<KiCadCalibrationEntry> calibrations) =
                DraftEditRules.Follow(before.Board, edited.Board, edited.Calibrations, follow);

            return new DraftPreparedEdit(before, edited with { Board = board, Calibrations = calibrations });
        }

        // ###########################################################################################
        // Applies an edit Prepare worked out, as one step - when the draft is still the content it was
        // worked out on. False, applying nothing, when the draft moved meanwhile: the caller works it
        // out again on the draft as it is (ApplyWithFollow). UI thread, like Apply.
        // ###########################################################################################
        public bool ApplyPrepared(string description, DraftPreparedEdit prepared, object? origin = null)
        {
            ArgumentNullException.ThrowIfNull(prepared);

            this.ChangeComing?.Invoke(this, new DraftSessionChange(DraftSessionChangeKind.Edit, description ?? string.Empty, origin));

            return this.Commit(description, prepared, origin);
        }

        private bool Commit(string? description, DraftPreparedEdit prepared, object? origin)
        {
            lock (this.thisGate)
            {
                if (!ReferenceEquals(this.thisContent, prepared.Before))
                {
                    return false;
                }

                this.thisRedo.Clear();
                this.thisUndo.Add(new Step(prepared.Before, description ?? string.Empty, origin));

                if (this.thisUndo.Count > DraftSession.MaxSteps)
                {
                    this.thisUndo.RemoveAt(0);
                }

                this.thisContent = prepared.After;
                this.thisLastChange = this.thisClock();
            }

            this.Raise(new DraftSessionChange(DraftSessionChangeKind.Edit, description ?? string.Empty, origin));

            return true;
        }

        public bool CanUndo
        {
            get
            {
                lock (this.thisGate)
                    return this.thisUndo.Count > 0;
            }
        }

        public bool CanRedo
        {
            get
            {
                lock (this.thisGate)
                    return this.thisRedo.Count > 0;
            }
        }

        // What an undo or a redo would take back or do again - null with nothing to take.
        public string? UndoDescription
        {
            get
            {
                lock (this.thisGate)
                    return this.thisUndo.Count > 0 ? this.thisUndo[^1].Description : null;
            }
        }

        public string? RedoDescription
        {
            get
            {
                lock (this.thisGate)
                    return this.thisRedo.Count > 0 ? this.thisRedo[^1].Description : null;
            }
        }

        // Takes back the latest change, whichever editor made it. False with nothing to undo.
        public bool Undo() => this.Move(this.thisUndo, this.thisRedo, DraftSessionChangeKind.Undo);

        public bool Redo() => this.Move(this.thisRedo, this.thisUndo, DraftSessionChangeKind.Redo);

        private bool Move(List<Step> from, List<Step> to, DraftSessionChangeKind kind)
        {
            Step step;

            Step? coming;

            lock (this.thisGate)
            {
                coming = from.Count == 0 ? null : from[^1];
            }

            if (coming is null)
            {
                return false;
            }

            this.ChangeComing?.Invoke(this, new DraftSessionChange(kind, coming.Description, coming.Origin));

            lock (this.thisGate)
            {
                if (from.Count == 0)
                {
                    return false;
                }

                step = from[^1];
                from.RemoveAt(from.Count - 1);

                // The state being left goes on the other stack, so the opposite command returns to it.
                to.Add(step with { Content = this.thisContent });

                this.thisContent = step.Content;
                this.thisLastChange = this.thisClock();
            }

            this.Raise(new DraftSessionChange(kind, step.Description, step.Origin));

            return true;
        }

        // ###########################################################################################
        // TAKES BACK a change whose save could not be made - the step `origin` made. Not redoable: it
        // never happened as far as the editor that made it is concerned ("could not be saved"). True
        // when it took it back; false when no step of `origin`'s is in the history.
        //
        // *** ALSO WHEN ANOTHER CHANGE CAME AFTER IT (code review, 2026-10-25). *** The component editor
        // is a window of its own, so the Drafts tab's table can push a change while the editor's save
        // runs - and the editor then said "could not be saved" while its change stayed, to be saved by
        // the table's next save. A step that is not the latest is taken back by the three-way merge
        // "Bring my draft up to date" uses (DraftRebase - base: the draft right after the change,
        // draft: as it is now, update: right before the change): what nobody touched since goes back,
        // and a row changed since keeps its later change. Every later snapshot in the history loses it
        // the same way, so an undo cannot bring it back.
        // ###########################################################################################
        public bool TakeBack(object origin)
        {
            ArgumentNullException.ThrowIfNull(origin);

            Step? latest;
            bool earlier;

            lock (this.thisGate)
            {
                latest = this.thisUndo.Count == 0 ? null : this.thisUndo[^1];
                earlier = this.thisUndo.Take(Math.Max(0, this.thisUndo.Count - 1)).Any(step => ReferenceEquals(step.Origin, origin));
            }

            if (latest is null || !ReferenceEquals(latest.Origin, origin))
            {
                return earlier && this.TakeBackEarlier(origin);
            }

            this.ChangeComing?.Invoke(this, new DraftSessionChange(DraftSessionChangeKind.TakenBack, latest.Description, origin));

            bool stillLatest;

            lock (this.thisGate)
            {
                stillLatest = this.thisUndo.Count > 0 && ReferenceEquals(this.thisUndo[^1], latest);

                if (stillLatest)
                {
                    this.thisUndo.RemoveAt(this.thisUndo.Count - 1);
                    this.thisContent = latest.Content;
                    this.thisLastChange = this.thisClock();
                }
            }

            // *** A CHANGE HANDED OVER JUST NOW (code review, 2026-10-26). *** An editor holding one hands
            // it over as the take-back is said (the table's open cell, the pane's typed box), so the step
            // is no longer the latest: taken back from under it, as from under any later change. It
            // gave up here, and the editor said "could not be saved" while its change stayed.
            if (!stillLatest)
            {
                return this.TakeBackEarlier(origin);
            }

            this.Raise(new DraftSessionChange(DraftSessionChangeKind.TakenBack, latest.Description, origin));

            return true;
        }

        // The step `origin` made, taken back from under the steps that came after it - see TakeBack.
        private bool TakeBackEarlier(object origin)
        {
            string description;

            lock (this.thisGate)
            {
                int index = this.thisUndo.FindLastIndex(step => ReferenceEquals(step.Origin, origin));

                if (index < 0)
                {
                    return false;
                }

                description = this.thisUndo[index].Description;
            }

            this.ChangeComing?.Invoke(this, new DraftSessionChange(DraftSessionChangeKind.TakenBack, description, origin));

            lock (this.thisGate)
            {
                int index = this.thisUndo.FindLastIndex(step => ReferenceEquals(step.Origin, origin));

                if (index < 0)
                {
                    return false;
                }

                BoardContent before = this.thisUndo[index].Content;
                BoardContent after = index + 1 < this.thisUndo.Count ? this.thisUndo[index + 1].Content : this.thisContent;

                for (int later = index + 1; later < this.thisUndo.Count; later++)
                {
                    this.thisUndo[later] = this.thisUndo[later] with { Content = DraftSession.WithoutChange(this.thisUndo[later].Content, after, before) };
                }

                for (int redo = 0; redo < this.thisRedo.Count; redo++)
                {
                    this.thisRedo[redo] = this.thisRedo[redo] with { Content = DraftSession.WithoutChange(this.thisRedo[redo].Content, after, before) };
                }

                this.thisUndo.RemoveAt(index);
                this.thisContent = DraftSession.WithoutChange(this.thisContent, after, before);
                this.thisLastChange = this.thisClock();
            }

            this.Raise(new DraftSessionChange(DraftSessionChangeKind.TakenBack, description, origin));

            return true;
        }

        // ###########################################################################################
        // `content` without the change that turned `before` into `after`: rows the changes since left
        // as that change made them go back, and rows changed since keep what they are now - DraftRebase's
        // three-way merge with `after` as the base, `content` as the draft and `before` as the update.
        // ###########################################################################################
        private static BoardContent WithoutChange(BoardContent content, BoardContent after, BoardContent before)
        {
            if (ReferenceEquals(content, after))
            {
                return before;
            }

            DraftRebaseResult merged = DraftRebase.Merge(
                after.Board,
                content.Board,
                before.Board,
                after.Calibrations,
                content.Calibrations,
                before.Calibrations);

            return content with { Board = merged.Board, Calibrations = merged.Calibrations };
        }

        // ###########################################################################################
        // THE FILE'S VERSION, after a save refused because the file changed outside CRT: read again,
        // as ONE undoable step - what was on screen is one Ctrl+Z away. Saved afterwards: the file
        // holds what is shown. False, changing nothing, when the file cannot be read.
        // ###########################################################################################
        public bool ReloadFromFile(string description)
        {
            this.ChangeComing?.Invoke(this, new DraftSessionChange(DraftSessionChangeKind.ReadFromFile, description ?? string.Empty, null));

            string fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);
            BoardContent? content = this.ReadBoardFile();

            if (content is null || fingerprint.Length == 0)
            {
                return false;
            }

            lock (this.thisGate)
            {
                this.thisRedo.Clear();
                this.thisUndo.Add(new Step(this.thisContent, description ?? string.Empty, Origin: null));

                if (this.thisUndo.Count > DraftSession.MaxSteps)
                {
                    this.thisUndo.RemoveAt(0);
                }

                this.thisContent = content;
                this.thisSavedContent = content;
                this.thisFingerprint = fingerprint;
                this.thisSettledStamp = null;
                this.thisFailure = null;
                this.thisFailedContent = null;
                this.thisLastChange = this.thisClock();
            }

            this.Raise(new DraftSessionChange(DraftSessionChangeKind.ReadFromFile, description ?? string.Empty, null));

            return true;
        }

        // ###########################################################################################
        // THE FILE WRITTEN BY SOMETHING ELSE, while nothing of this session is waiting to be saved -
        // another part of CRT that still writes the draft's file itself: read again, as one undoable
        // step (ReloadFromFile). Nothing when the file is what this session read or wrote, and
        // nothing either while a change of its own waits: that save refuses (ChangedOnDisk) and the
        // host asks which version to keep - neither is thrown away unasked. True when it read.
        // ###########################################################################################
        public bool ReadFileIfChanged(string description)
        {
            // A change held by an editor is handed over first - then it waits, and nothing is read.
            this.CollectHeldChanges(description);

            if (!this.FileMoved())
            {
                return false;
            }

            lock (this.thisGate)
            {
                if (this.thisSaving || !ReferenceEquals(this.thisContent, this.thisSavedContent))
                {
                    return false;
                }
            }

            return this.ReloadFromFile(description);
        }

        // ###########################################################################################
        // Whether the file is no longer what this session read or wrote - its fingerprint, so any
        // thread (an editor asks off the UI thread before it reads the file again there with
        // ReadFileIfChanged, which hashes it once more: a board file of a few megabytes is hashed
        // on the UI thread only when it really changed).
        // ###########################################################################################
        public bool FileChangedOutside() => this.FileMoved();

        // How often the board file was hashed to answer FileMoved - for a test.
        internal int FingerprintsTakenForTests => Volatile.Read(ref this.thisFingerprintsTaken);

        // ###########################################################################################
        // Whether the file's fingerprint differs from thisFingerprint (one that cannot be read does
        // not). *** ITS STAMP FIRST (code review, 2026-10-26). *** The table asks at every switch to the
        // Contribute tab and the pane at every component opened, on the UI thread, and the board file
        // was hashed each time. Once it was seen unchanged with a stamp that had SETTLED - two writes
        // inside one clock tick share a stamp (StampSettling, git's "racily clean") - the same stamp is
        // the same file, and the hash is skipped. Any new fingerprint forgets the stamp.
        // ###########################################################################################
        //
        // *** NOT WHILE A WRITE OF ITS OWN RUNS (code review, 2026-10-26). *** Between a save's write and
        // its taking the new file as its own, the new bytes looked like an outside change against the
        // old fingerprint, and ReadFileIfChanged read the file back as a step - Redo lost, "updated from
        // outside" said with nothing outside. Save and WriteFileDirectly hold thisWriting; while another
        // thread does, nothing is judged - the next ask is - and while this one asks, no write starts.
        private bool FileMoved()
        {
            if (!Monitor.TryEnter(this.thisWriting))
            {
                return false;
            }

            try
            {
                return this.FileMovedWhileNothingWrites();
            }
            finally
            {
                Monitor.Exit(this.thisWriting);
            }
        }

        private bool FileMovedWhileNothingWrites()
        {
            string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, this.ExcelDataFile);
            string stamp = BoardDataReader.FileStamp(boardFile);
            string known;

            lock (this.thisGate)
            {
                if (this.thisSettledStamp is not null && string.Equals(stamp, this.thisSettledStamp, StringComparison.Ordinal))
                {
                    return false;
                }

                known = this.thisFingerprint;
            }

            Interlocked.Increment(ref this.thisFingerprintsTaken);
            string fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);

            if (fingerprint.Length == 0 || !string.Equals(fingerprint, known, StringComparison.Ordinal))
            {
                return fingerprint.Length > 0;
            }

            // Unchanged - and its stamp trusted from now on, once settled and still the one hashed.
            if (StampSettling.IsSettled(File.GetLastWriteTimeUtc(boardFile)) &&
                string.Equals(BoardDataReader.FileStamp(boardFile), stamp, StringComparison.Ordinal))
            {
                lock (this.thisGate)
                {
                    if (string.Equals(this.thisFingerprint, known, StringComparison.Ordinal))
                    {
                        this.thisSettledStamp = stamp;
                    }
                }
            }

            return false;
        }

        // ---- saving ----------------------------------------------------------------------------

        // Whether the file holds what is shown - false between a change and its save.
        public bool HasUnsavedChanges
        {
            get
            {
                lock (this.thisGate)
                    return !ReferenceEquals(this.thisContent, this.thisSavedContent);
            }
        }

        public DraftSaveState SaveState
        {
            get
            {
                lock (this.thisGate)
                {
                    if (this.thisSaving)
                        return DraftSaveState.Saving;

                    if (ReferenceEquals(this.thisContent, this.thisSavedContent))
                        return DraftSaveState.Saved;

                    return this.thisFailure is null ? DraftSaveState.Pending : DraftSaveState.Failed;
                }
            }
        }

        // Why the last save failed while SaveState is Failed - WriteFailed or ChangedOnDisk.
        public DraftWorkbookEditOutcome? SaveFailure
        {
            get
            {
                lock (this.thisGate)
                    return this.thisFailure;
            }
        }

        // ###########################################################################################
        // Whether a save should run now: something unsaved, nothing being written, the last change
        // SaveDelay old - and, after a failed save, a change made since (or the host's Retry).
        // ###########################################################################################
        public bool IsSaveDue(DateTimeOffset now)
        {
            lock (this.thisGate)
            {
                return !this.thisSaving &&
                       !ReferenceEquals(this.thisContent, this.thisSavedContent) &&
                       !ReferenceEquals(this.thisContent, this.thisFailedContent) &&
                       now - this.thisLastChange >= DraftSession.SaveDelay;
            }
        }

        // The same, now by the session's own clock - what the host's timer asks.
        public bool IsSaveDueNow => this.IsSaveDue(this.thisClock());

        // "Try again": the failed content is due as soon as the host looks.
        public void Retry()
        {
            lock (this.thisGate)
            {
                this.thisFailedContent = null;
                this.thisLastChange = DateTimeOffset.MinValue;
            }
        }

        // Runs inside Save after it took the content and before it writes - for a test to make a
        // change land WHILE a save runs, which a test cannot time otherwise.
        internal Action? WhileWritingForTests { get; set; }

        // Runs inside Save after the file is written and before the session takes it as its own - for
        // a test to ask about the file at that moment (FileMoved).
        internal Action? AfterWritingForTests { get; set; }

        // ###########################################################################################
        // Writes the draft as it is now, the board file replaced whole (DraftWorkbookStore - new files
        // lying in the draft folder taken into its store on the way). Refused when the file changed
        // since this session read or wrote it, unless `overwriteOutsideChange` - the host's "keep the
        // version on screen". Saved straight away when there is nothing to write.
        //
        // A change made while it writes is not lost: Saved covers the content written, and the newer
        // content stays unsaved, due again after SaveDelay.
        //
        // `evenUnchanged` writes the board file with nothing changed in it - so the new files lying in
        // the draft folder (a KiCad project just imported) are taken in and listed (DraftFiles).
        // ###########################################################################################
        public DraftWorkbookEditOutcome Save(bool overwriteOutsideChange = false, bool evenUnchanged = false)
        {
            lock (this.thisWriting)
            {
                BoardContent content;
                string expected;

                lock (this.thisGate)
                {
                    content = this.thisContent;

                    if (ReferenceEquals(content, this.thisSavedContent) && !evenUnchanged)
                    {
                        return DraftWorkbookEditOutcome.Saved;
                    }

                    expected = this.thisFingerprint;
                    this.thisSaving = true;
                }

                this.SaveStateChanged?.Invoke(this, EventArgs.Empty);

                // Failed until the write says otherwise - so an exception nobody expected still ends
                // the save as a failure, never with the session stuck in Saving (code review,
                // 2026-10-25: nothing was ever due again, and nothing said so).
                DraftWorkbookEditOutcome outcome = DraftWorkbookEditOutcome.WriteFailed;
                string written = expected;

                try
                {
                    this.WhileWritingForTests?.Invoke();

                    if (overwriteOutsideChange)
                    {
                        expected = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);
                    }

                    outcome = DraftWorkbookStore.ReplaceIfUnchanged(
                        this.DraftsRoot,
                        this.ExcelDataFile,
                        expected,
                        content.Board,
                        content.Calibrations,
                        content.Renames,
                        memory: this.thisFiles);

                    if (outcome == DraftWorkbookEditOutcome.Saved)
                    {
                        written = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);
                    }

                    this.AfterWritingForTests?.Invoke();
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
                {
                    CrtLog.Warning($"Could not save the draft for [{this.ExcelDataFile}] - [{ex.Message}]");
                    outcome = DraftWorkbookEditOutcome.WriteFailed;
                }
                finally
                {
                    lock (this.thisGate)
                    {
                        this.thisSaving = false;

                        if (outcome == DraftWorkbookEditOutcome.Saved)
                        {
                            this.thisSavedContent = content;
                            this.thisFingerprint = written;
                            this.thisSettledStamp = null;
                            this.thisFailure = null;
                            this.thisFailedContent = null;
                        }
                        else
                        {
                            // The content that failed waits for a change or a Retry; a newer one that
                            // landed while it was written is due as usual.
                            this.thisFailure = outcome;
                            this.thisFailedContent = content;
                        }
                    }

                    this.SaveStateChanged?.Invoke(this, EventArgs.Empty);
                }

                return outcome;
            }
        }

        // ###########################################################################################
        // Save, OFF THE CALLING THREAD - every host's save, the one way (code review, 2026-10-26: the
        // table, the pane, the header and the tab each had a copy of "run it on the pool, and a save
        // that throws is a failed save"). Never throws: a save stopped by an exception nobody expected
        // ends as WriteFailed (Save leaves the session Failed), logged here. One write at a time, as
        // Save.
        // ###########################################################################################
        public async Task<DraftWorkbookEditOutcome> SaveAsync(bool overwriteOutsideChange = false, bool evenUnchanged = false)
        {
            try
            {
                return await Task.Run(() => this.Save(overwriteOutsideChange, evenUnchanged)).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Saving the draft for [{this.ExcelDataFile}] failed unexpectedly - [{ex}]");
                return DraftWorkbookEditOutcome.WriteFailed;
            }
        }

        // ###########################################################################################
        // A WRITE OF THE DRAFT'S FILE THAT IS MORE THAN ITS CONTENT - the board's file list
        // (DraftWorkbookStore.RemoveFiles), a draft brought up to date together with its base and its
        // marker (DraftRebaser) - run under the session's own lock: whatever waits is saved first (a
        // save that cannot be made stops it, and `write` is not run), then `write`, then the file is
        // read back as what the session holds and has saved. `clearHistory` drops the undo steps -
        // after bringing a draft up to date they would put back rows measured against the old base.
        // `succeeded` says whether `write` wrote: only then is the file read back. Runs anywhere;
        // Changed is raised on the calling thread. Its callers are windows that hold CRT's main window
        // (modal), so no editor changes the content while it runs - a change made then would be
        // replaced by the file read back.
        // ###########################################################################################
        public T WriteFileDirectly<T>(Func<T> write, Func<T, bool> succeeded, string description, bool clearHistory)
        {
            // What an editor still holds was handed over before this left the UI thread
            // (CollectHeldChanges) - ChangeComing belongs to the UI thread, and this need not run there.
            ArgumentNullException.ThrowIfNull(write);
            ArgumentNullException.ThrowIfNull(succeeded);

            lock (this.thisWriting)
            {
                if (this.HasUnsavedChanges && this.Save() != DraftWorkbookEditOutcome.Saved)
                {
                    throw new DraftNotSavedException(this.SaveFailure ?? DraftWorkbookEditOutcome.WriteFailed);
                }

                T result = write();

                if (!succeeded(result))
                {
                    return result;
                }

                // With the history gone, nothing can cite again what it remembered of the list.
                if (clearHistory)
                {
                    this.thisFiles.Forget();
                }

                string fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);
                BoardContent? content = this.ReadBoardFile();

                if (content is null || fingerprint.Length == 0)
                {
                    return result;
                }

                lock (this.thisGate)
                {
                    if (clearHistory)
                    {
                        this.thisUndo.Clear();
                        this.thisRedo.Clear();
                    }

                    this.thisContent = content;
                    this.thisSavedContent = content;
                    this.thisFingerprint = fingerprint;
                    this.thisSettledStamp = null;
                    this.thisFailure = null;
                    this.thisFailedContent = null;
                }

                this.Raise(new DraftSessionChange(DraftSessionChangeKind.ReadFromFile, description ?? string.Empty, null));

                return result;
            }
        }

        // ###########################################################################################
        // Every editor holding a change not handed over yet hands it over now (ChangeComing), changing
        // nothing else - on the UI thread, before a write of the file that runs off it
        // (WriteFileDirectly), so the write saves it first rather than read the file back over it.
        // ###########################################################################################
        public void CollectHeldChanges(string description) =>
            this.ChangeComing?.Invoke(this, new DraftSessionChange(DraftSessionChangeKind.ReadFromFile, description ?? string.Empty, null));

        private void Raise(DraftSessionChange change)
        {
            this.Changed?.Invoke(this, change);
            this.SaveStateChanged?.Invoke(this, EventArgs.Empty);
        }

        // One step of the history: the content to go back to, what the change was, who made it.
        private sealed record Step(BoardContent Content, string Description, object? Origin);
    }

    public enum DraftSessionChangeKind
    {
        Edit,
        Undo,
        Redo,
        ReadFromFile,

        // A change whose save could not be made, taken back out (DraftSession.TakeBack).
        TakenBack
    }

    // ###########################################################################################
    // A write of the draft's file could not run because what the session held could not be saved
    // first (DraftSession.WriteFileDirectly) - `Failure` says why.
    // ###########################################################################################
    public sealed class DraftNotSavedException(DraftWorkbookEditOutcome failure)
        : InvalidOperationException($"The draft's changes could not be saved first ({failure}).")
    {
        public DraftWorkbookEditOutcome Failure { get; } = failure;
    }

    // An edit worked out off the UI thread (DraftSession.Prepare): the content it was worked out on,
    // and the content it makes of it, what follows the edit followed.
    public sealed record DraftPreparedEdit(BoardContent Before, BoardContent After);

    // What changed a draft session: the kind, the step's description, and the editor that made the
    // edit (null for undo and redo of nobody's, and for the file read again).
    public sealed record DraftSessionChange(DraftSessionChangeKind Kind, string Description, object? Origin);

    // Where a draft stands against its file: Saved (the file holds what is shown), Pending (a change
    // waits for its save), Saving, or Failed (the last save could not be made - DraftSession.SaveFailure
    // says why).
    public enum DraftSaveState
    {
        Saved,
        Pending,
        Saving,
        Failed
    }
}
