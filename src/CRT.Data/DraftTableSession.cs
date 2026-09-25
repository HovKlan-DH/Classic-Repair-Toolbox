using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE SITTING IN THE TABLE EDITOR: the draft it was read from, the table, and the fingerprint
    // that decides whether a save is still safe (owner request, 2026-09-24).
    //
    // *** THE FINGERPRINT IS TAKEN BEFORE THE READ, NOT AFTER. *** If Excel saves between the two,
    // the table holds the newer content but the older fingerprint - so the first save refuses and
    // offers a reload, which is harmless. The other order would pair the older content with the
    // newer fingerprint, and that save would go through and silently undo the Excel edit.
    //
    // A successful save re-fingerprints, so the next save in the same sitting compares against
    // what THIS save wrote rather than refusing forever against the file as it was opened.
    // ###########################################################################################
    public sealed class DraftTableSession
    {
        private DraftTableSession(
            string draftsRoot,
            string excelDataFile,
            string fingerprint,
            BoardTableDocument document)
        {
            this.DraftsRoot = draftsRoot;
            this.ExcelDataFile = excelDataFile;
            this.Fingerprint = fingerprint;
            this.Document = document;
        }

        public string DraftsRoot { get; }

        public string ExcelDataFile { get; }

        public BoardTableDocument Document { get; }

        public string Fingerprint { get; private set; }

        // ###########################################################################################
        // Opens the draft's table, or null when there is no readable draft. `published` is the
        // board to colour differences against - null for a system with no published copy.
        // ###########################################################################################
        public static DraftTableSession? Open(string draftsRoot, string excelDataFile, BoardData? published)
        {
            string fingerprint = DraftWorkbookStore.Fingerprint(draftsRoot, excelDataFile);
            if (fingerprint.Length == 0)
            {
                return null;
            }

            BoardData? draft = DraftWorkbookStore.LoadDraftBoard(draftsRoot, excelDataFile);
            if (draft is null)
            {
                return null;
            }

            return new DraftTableSession(
                draftsRoot,
                excelDataFile,
                fingerprint,
                BoardTableDocument.Create(published, draft));
        }

        // ###########################################################################################
        // Writes the table into the draft, unless the workbook changed since it was read.
        //
        // Every sheet is brought up to date first: a cell committed a moment ago may not have been
        // re-paired yet (the UI refreshes just after an edit), and the rows saved should be the
        // ones on screen.
        // ###########################################################################################
        public DraftWorkbookEditOutcome Save()
        {
            foreach (BoardTableSheet sheet in this.Document.Sheets)
            {
                if (sheet.NeedsRefresh)
                {
                    sheet.Refresh();
                }
            }

            DraftWorkbookEditOutcome outcome = DraftWorkbookStore.EditIfUnchanged(
                this.DraftsRoot,
                this.ExcelDataFile,
                this.Fingerprint,
                current => this.Document.ApplyTo(current));

            if (outcome == DraftWorkbookEditOutcome.Saved)
            {
                this.Fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);
                this.Document.MarkSaved();
            }

            return outcome;
        }

        // True when the workbook on disk is no longer the one this table was read from - so the
        // UI can say "changed outside the table" before the contributor has typed anything.
        public bool HasChangedOnDisk() =>
            !string.Equals(
                DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile),
                this.Fingerprint,
                StringComparison.Ordinal);

        // ###########################################################################################
        // WATCHING THE DRAFT FILE WHILE THE TABLE IS OPEN (owner request, 2026-09-24): what is
        // true of the file right now - changed since the table read it, gone, or open in Excel (or
        // LibreOffice). The table asks every couple of seconds while it is on screen, so an edit
        // saved in Excel is noticed straight away rather than when the table's own save is refused.
        //
        // *** CHEAP WHEN NOTHING CHANGED. *** The file's time and size are compared with the last
        // check, and only when they differ is it hashed - so a check every two seconds costs one
        // stat call, not one read of the workbook.
        //
        // *** A FILE THAT CANNOT BE READ JUST NOW KEEPS THE LAST ANSWER. *** Excel saves by writing
        // a new file and swapping it in, and a check can land mid-swap; the next one decides.
        //
        // Changed is judged against the CURRENT fingerprint, which a successful save moves on - so
        // the table's own save is never mistaken for an outside one.
        // ###########################################################################################
        public DraftFileStatus CheckFile()
        {
            string workbook = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, this.ExcelDataFile);
            bool open = DraftTableSession.IsOpenInSpreadsheetProgram(workbook);

            var info = new FileInfo(workbook);
            if (workbook.Length == 0 || !info.Exists)
            {
                this.thisLastStamp = null;
                return new DraftFileStatus(ChangedSinceRead: false, IsMissing: true, IsOpenElsewhere: open, IsHeldOpenElsewhere: false);
            }

            (DateTime, long) stamp = (info.LastWriteTimeUtc, info.Length);
            if (this.thisLastStamp != stamp || !string.Equals(this.thisLastCheckedAgainst, this.Fingerprint, StringComparison.Ordinal))
            {
                string fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, this.ExcelDataFile);

                // Unreadable mid-save: keep the last answer, the next check decides.
                if (fingerprint.Length > 0)
                {
                    this.thisLastStamp = stamp;
                    this.thisLastCheckedAgainst = this.Fingerprint;
                    this.thisLastChanged = !string.Equals(fingerprint, this.Fingerprint, StringComparison.Ordinal);
                }
            }

            // Probed only while a lock file says a spreadsheet program has it - see IsHeldOpen.
            bool held = open && DraftTableSession.IsHeldOpen(workbook);

            return new DraftFileStatus(this.thisLastChanged, IsMissing: false, IsOpenElsewhere: open, IsHeldOpenElsewhere: held);
        }

        // ###########################################################################################
        // Whether another program holds the workbook open right now, so a save cannot replace it -
        // Excel does, for as long as it has the file open. Tried by opening it EXCLUSIVELY for a
        // moment: that fails while any other program has it open, and succeeds (and is closed at
        // once) when nothing does.
        //
        // *** WHY THIS, AND NOT THE LOCK FILE ALONE (2026-09-24). *** "Save changes" is switched off
        // while the draft is open in Excel (owner request), and a lock file alone would switch
        // it off for good after an Excel crash, which leaves its lock file behind. So the lock file
        // says "open in Excel", and this says whether that is still true.
        //
        // Only asked while a lock file exists, i.e. while Excel already has the file - never while
        // Excel might be opening it, where even a momentary exclusive open could get in its way. On
        // Linux and macOS file sharing is advisory, so this finds nothing held there; a save then
        // simply goes ahead, which is also what the operating system allows.
        // ###########################################################################################
        public static bool IsHeldOpen(string workbookPath)
        {
            try
            {
                using var probe = new FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.None);
                return false;
            }
            catch (FileNotFoundException)
            {
                return false;
            }
            catch (DirectoryNotFoundException)
            {
                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return true;
            }
        }

        private (DateTime, long)? thisLastStamp;
        private string thisLastCheckedAgainst = string.Empty;
        private bool thisLastChanged;

        // ###########################################################################################
        // Whether a spreadsheet program has the workbook open, going by the lock file it leaves
        // beside it while it does: Excel's hidden "~$" owner file (the whole name after "~$", or -
        // as Word and some Excel versions do - the name with its first two characters replaced) and
        // LibreOffice's ".~lock.<name>#". A program that crashed can leave one behind, which is why
        // the UI says the file APPEARS to be open.
        // ###########################################################################################
        public static bool IsOpenInSpreadsheetProgram(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                return false;
            }

            string? folder = Path.GetDirectoryName(workbookPath);
            if (string.IsNullOrEmpty(folder))
            {
                return false;
            }

            return DraftTableSession.LockFileNamesFor(Path.GetFileName(workbookPath))
                .Any(name => File.Exists(Path.Combine(folder, name)));
        }

        public static IReadOnlyList<string> LockFileNamesFor(string workbookFileName)
        {
            if (string.IsNullOrEmpty(workbookFileName))
            {
                return [];
            }

            var names = new List<string> { "~$" + workbookFileName, ".~lock." + workbookFileName + "#" };

            if (workbookFileName.Length > 2)
            {
                names.Add("~$" + workbookFileName[2..]);
            }

            return names;
        }
    }

    // What DraftTableSession.CheckFile found. IsMissing: the draft file is gone (discarded, say) -
    // not a change to warn about, the Drafts tab closes such a table itself. IsOpenElsewhere: a
    // spreadsheet program's lock file is there. IsHeldOpenElsewhere: and the workbook really is
    // held open, so a save cannot replace it now.
    public readonly record struct DraftFileStatus(bool ChangedSinceRead, bool IsMissing, bool IsOpenElsewhere, bool IsHeldOpenElsewhere);
}
