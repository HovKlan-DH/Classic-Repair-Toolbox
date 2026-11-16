using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE OPEN DRAFT SESSIONS - ONE PER DRAFT, whichever editor asks first (CRT 4.0 step 1). Every
    // editor of a board's draft asks here for its session, so they all hold the same one and see
    // each other's edits at once (DraftSession). The host keeps one of these, and its timer saves
    // whatever is due (Due).
    //
    // *** ONE PER DRAFT FOLDER, NOT PER IDENTITY STRING. *** A board's identity names its board file
    // ("<id>/board.json"), but a draft made before 4.0 still carries "<id>/Data X.xlsx" - and only the
    // folder of an identity matters (DataTreeLayout.BoardFileFor). Keyed by the identity, the two
    // would be two sessions of one draft, each saving over the other.
    // ###########################################################################################
    public sealed class DraftSessions
    {
        private readonly object thisGate = new();
        private readonly Dictionary<string, DraftSession> thisOpen = new(StringComparer.OrdinalIgnoreCase);
        private readonly Func<DateTimeOffset>? thisClock;

        // How many Close calls there have been - what a read asks to know whether one came while it ran.
        private long thisCloses;

        public DraftSessions(Func<DateTimeOffset>? clock = null)
        {
            this.thisClock = clock;
        }

        // Runs once a draft has been read, before its session is taken in - what a test does meanwhile.
        internal Action? AfterReadForTests { get; set; }

        // ###########################################################################################
        // WHOM A DRAFT'S CREDIT FOR ITS WORK NAMES - the name and contact CRT remembers, handed in by the app (2026-11-15). Set, the
        // sessions opened here follow their work with that credit (SubmissionCredit.Follow, against the draft's base -
        // DraftSession.BaseBoard), in the same step as each edit; null, none does - a test of the sessions themselves. Asked at each
        // edit, on the thread the edit is worked out on.
        // ###########################################################################################
        public Func<(string Name, string Contact)>? CreditIdentity { get; set; }

        // ###########################################################################################
        // A session's board AS PUBLISHED NOW - what the draft's work and the credits it sends are judged against, as Submit and
        // the server judge them (SubmissionCredit.Follow; code reviews, 2026-11-16). Null, or an answer of null (a board not
        // published, or its file unreadable), the draft's base. Asked at each edit, on the thread the edit is worked out on.
        // ###########################################################################################
        public Func<DraftSession, BoardData?>? PublishedBoard { get; set; }

        // The board a session's draft is judged against (above) - null for a new board.
        public BoardData? BoardToJudgeAgainst(DraftSession session)
        {
            ArgumentNullException.ThrowIfNull(session);

            return this.PublishedBoard?.Invoke(session) ?? session.BaseBoard();
        }

        private void FollowWorkWithCredit(DraftSession session) =>
            session.FollowUp = (before, after) =>
            {
                if (this.CreditIdentity is not { } who || this.BoardToJudgeAgainst(session) is not { } published)
                {
                    return after;
                }

                (string name, string contact) = who();

                return SubmissionCredit.Follow(published, before, after, name, contact);
            };

        // ###########################################################################################
        // The session of the draft of `excelDataFile` - the one already open, or a new one read from
        // its board file. Null when there is no such draft.
        //
        // *** THE BOARD FILE IS READ OUTSIDE THE LOCK (code review, 2026-11-02). *** Holding it for the
        // parse, the first open of a large draft froze the window even off the UI thread: the save timer
        // (Due) and every refresh (Find) wait for this lock on the UI thread. Reading changes nothing, so
        // two asking at once may both read, and the first to arrive is the one session. A Close meanwhile
        // read it for a draft that may have gone, so it is read again: taken in, it would be answered for
        // good, and its next save would write the folder back.
        // ###########################################################################################
        public DraftSession? Get(string draftsRoot, string excelDataFile)
        {
            string key = DraftSessions.KeyOf(draftsRoot, excelDataFile);

            if (key.Length == 0)
            {
                return null;
            }

            while (true)
            {
                long closes;

                lock (this.thisGate)
                {
                    if (this.thisOpen.TryGetValue(key, out DraftSession? open))
                    {
                        return open;
                    }

                    closes = this.thisCloses;
                }

                DraftSession? session = DraftSession.Open(draftsRoot, excelDataFile, this.thisClock);

                this.AfterReadForTests?.Invoke();

                lock (this.thisGate)
                {
                    if (this.thisOpen.TryGetValue(key, out DraftSession? open))
                    {
                        return open;
                    }

                    if (closes != this.thisCloses)
                    {
                        continue;
                    }

                    if (session is null)
                    {
                        return null;
                    }

                    this.FollowWorkWithCredit(session);
                    this.thisOpen[key] = session;
                }

                this.ReadWhatTheCreditFollowsBy(session);

                return session;
            }
        }

        // ###########################################################################################
        // *** WHAT THE CREDIT FOLLOWS THE WORK BY IS READ AS THE SESSION OPENS (code review, 2026-11-16). *** The follow-up runs
        // inside every edit, some of them applied on the UI thread: the first edit of a draft parsed its base and the board as
        // published there, a visible hitch on a large board. A session is opened where its board file is read - off the UI
        // thread wherever an editor opens one - so both are read here, outside the sessions' lock, and an edit finds them kept.
        // ###########################################################################################
        private void ReadWhatTheCreditFollowsBy(DraftSession session)
        {
            if (this.CreditIdentity is null)
            {
                return;
            }

            _ = session.BaseBoard();
            _ = this.PublishedBoard?.Invoke(session);
        }

        // The session of that draft when one is open - opening none.
        public DraftSession? Find(string draftsRoot, string excelDataFile)
        {
            string key = DraftSessions.KeyOf(draftsRoot, excelDataFile);

            lock (this.thisGate)
                return key.Length > 0 && this.thisOpen.TryGetValue(key, out DraftSession? open) ? open : null;
        }

        public IReadOnlyList<DraftSession> Open
        {
            get
            {
                lock (this.thisGate)
                    return this.thisOpen.Values.ToList();
            }
        }

        // The sessions with a save due at `now` (DraftSession.IsSaveDue) - what the Contribute tab's
        // timer saves of those no editor holds (TabDrafts.SaveDueDraftsNoEditorHoldsAsync; the table
        // and the component pane save their own, by their own timers).
        public IReadOnlyList<DraftSession> Due(DateTimeOffset now) =>
            this.Open.Where(session => session.IsSaveDue(now)).ToList();

        // ###########################################################################################
        // Forgets the session of a draft - after it is discarded, or retired once published. Its
        // unsaved changes are the caller's to save or throw away first; a later Get reads the file.
        // ###########################################################################################
        public void Close(string draftsRoot, string excelDataFile)
        {
            string key = DraftSessions.KeyOf(draftsRoot, excelDataFile);

            lock (this.thisGate)
            {
                this.thisOpen.Remove(key);
                this.thisCloses++;
            }
        }

        private static string KeyOf(string draftsRoot, string excelDataFile)
        {
            string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile);

            return folder.Length == 0 ? string.Empty : Path.GetFullPath(folder);
        }
    }
}
