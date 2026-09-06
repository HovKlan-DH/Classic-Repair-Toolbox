using Avalonia;
using CRT;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One repair job a user is tracking against a board. Persisted as its own subfolder's
    // index.json inside the "Workbooks" folder - never synced and never part of the online "Data"
    // folder. Everything that ever belongs to this workbook (entries, photos, files) lives inside
    // that same subfolder, so deleting the folder deletes the workbook entirely.
    // ###########################################################################################
    public sealed class WorkbookRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("boardKey")] public string BoardKey { get; set; } = string.Empty;
        [JsonPropertyName("title")] public string Title { get; set; } = string.Empty;
        [JsonPropertyName("note")] public string Note { get; set; } = string.Empty;
        [JsonPropertyName("status")] public string Status { get; set; } = "Open";
        [JsonPropertyName("startDate")] public DateTime StartDate { get; set; }

        // ###########################################################################################
        // When this workbook last became Closed - i.e. the date the last outstanding worklog in it
        // was closed, so the top bar can report a finished job by when it FINISHED rather than by
        // when it started.
        //
        // Written only by RecomputeWorkbookStatus, and only on the transition into Closed: a
        // workbook that reopens (a new worklog, or an existing one set back to Open) keeps the
        // stale value, which is harmless because nothing reads it while Status is Open, and closing
        // it again overwrites it with the new date. StartDate is deliberately never touched by any
        // of this - a job's start is a fact about the past.
        //
        // Null means "never closed", which is what every workbook written before this field existed
        // deserializes to, INCLUDING ones already Closed. Nullable rather than DateTime for exactly
        // that reason: default(DateTime) is 0001-01-01, which is a date, and a top bar reading
        // "ended 0001-January-01" is worse than falling back to the start date. Every reader must
        // handle the null - see Main.RefreshWorklogBar, which falls back rather than inventing one.
        // There is deliberately no migration; an old closed workbook shows its start date until the
        // next time it actually closes.
        // ###########################################################################################
        [JsonPropertyName("endDate")] public DateTime? EndDate { get; set; }

        // How many worklogs this workbook holds.
        //
        // "worklog", not "entry", in the JSON AND in the property. This file is read by users, and
        // the app calls these worklogs everywhere one is visible - "entry" is internal vocabulary
        // that must not leak into stored data any more than it may leak into the UI. It was
        // "entryCount"; there is deliberately no migration, so a workbook written by an older build
        // reads this as 0 until it is next recomputed (which RecomputeWorkbookStatus does on any
        // worklog change).
        [JsonPropertyName("worklogCount")] public int WorklogCount { get; set; }

        // ###########################################################################################
        // The user's chosen top-to-bottom order for the schematic previews in the Workbooks tab's
        // board pane, as schematic NAMES. Empty means "never rearranged", which is every workbook
        // written before this existed - the pane then falls back to its alphabetical grouping, so an
        // absent value needs no migration and changes nothing.
        //
        // The ordering rules (why names rather than indices, where a newly-appearing schematic goes,
        // what happens to a stored name whose worklogs were all deleted) live with the logic in
        // WorkbookSchematicOrder rather than here - this is only where it is persisted.
        // ###########################################################################################
        [JsonPropertyName("schematicOrder")] public List<string> SchematicOrder { get; set; } = new();

        // ###########################################################################################
        // The highest worklog id ever HANDED OUT for this workbook - not the highest currently on
        // disk. The next worklog takes this plus one, so deleting one never lets its number come
        // back.
        //
        // Zero means "not recorded yet", which is what every workbook written before this field
        // existed deserializes to. It is NOT the same as "no worklogs have been created": such a
        // workbook may well have worklogs, so the counter is seeded from the highest id actually on
        // disk the first time it is needed - see AllocateNextEntryId.
        //
        // "worklog" rather than "entry" in the stored name for the reason WorklogCount above gives.
        // It was "lastEntryId", and is likewise not migrated - an older workbook reads 0 here, which
        // the walk-past-what-exists floor already handles safely.
        // ###########################################################################################
        [JsonPropertyName("lastWorklogId")] public int LastWorklogId { get; set; }
    }

    // ###########################################################################################
    // One worklog entry: the drawn area on a schematic, its headline/comment, category
    // (Note/Cosmetic/Issue), resolution state (Open/Closed) and the board labels of the
    // components marked in scope. Persisted as its OWN index.json inside its own
    // "worklog_<id>" folder, beside the photos and files belonging to it - see the WorklogManager
    // class header for why there is no file listing worklogs.
    //
    // Title is the entry's one-line headline and Description its longer comment. Both are written
    // when the entry is created and both stay editable afterwards - neither is legacy, and dropping
    // either as "redundant" would discard real user data.
    // Links/Comments/WorkDoneItems/Photos/Files are the full editor's own sub-lists - see
    // their own record types below. Photo/file bytes themselves are not stored here, only their
    // metadata; the files live in the entry's own "worklog_<id>" subfolder under the workbook.
    // ###########################################################################################
    public sealed class WorklogEntryRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("schematicName")] public string SchematicName { get; set; } = string.Empty;
        [JsonPropertyName("areaX")] public double AreaX { get; set; }
        [JsonPropertyName("areaY")] public double AreaY { get; set; }
        [JsonPropertyName("areaWidth")] public double AreaWidth { get; set; }
        [JsonPropertyName("areaHeight")] public double AreaHeight { get; set; }
        [JsonPropertyName("title")] public string Title { get; set; } = string.Empty;
        [JsonPropertyName("description")] public string Description { get; set; } = string.Empty;
        [JsonPropertyName("category")] public string Category { get; set; } = "Note";
        [JsonPropertyName("state")] public string State { get; set; } = "Open";
        [JsonPropertyName("componentLabels")] public List<string> ComponentLabels { get; set; } = new();

        // Whether the entry's coloured area is drawn on the Schematics tab. Defaults to true, which
        // is also what an entry written by an older build gets: the property is simply absent from
        // its JSON, so the initialiser stands and every existing worklog keeps showing its area.
        // A default of false would silently blank the board for anyone upgrading.
        //
        // An entry with this off still shows its "#N" pill - parked in the top-right corner of the
        // schematic panel rather than on the board (see ParkedBadgeGeometry). Hiding the pill too
        // would make the entry invisible and unreachable from the board.
        [JsonPropertyName("showMarkedArea")] public bool ShowMarkedArea { get; set; } = true;

        // Which of the in-scope components the user has ticked off as done - the "Mark components
        // completed" checklist, for tracking progress through a job like "replace every capacitor".
        //
        // Always a SUBSET of ComponentLabels: the completed list offers exactly the components that
        // are in scope, so a label dropped from the scope is dropped from here too rather than
        // lingering as a completed component the entry no longer covers. Empty by default, and for
        // every entry written before this field existed - a new component is work still to do, so
        // "not started" is the only honest starting point.
        [JsonPropertyName("completedComponentLabels")] public List<string> CompletedComponentLabels { get; set; } = new();

        // Which of the editor's list sections (Links, Work done, Comments, Components in scope,
        // Components completed, Photos, Files) the user has folded away, keyed by section name.
        //
        // A map of only the COLLAPSED sections rather than a flag per list: sections are expanded by
        // default, so an absent key means "open" and an entry written before this existed opens with
        // everything showing. It also means adding an eighth list later needs no schema change and
        // no migration - the new section simply has no key yet.
        [JsonPropertyName("collapsedSections")] public List<string> CollapsedSections { get; set; } = new();
        [JsonPropertyName("createdDate")] public DateTime CreatedDate { get; set; }

        [JsonPropertyName("links")] public List<WorklogLinkRecord> Links { get; set; } = new();
        [JsonPropertyName("comments")] public List<WorklogCommentRecord> Comments { get; set; } = new();
        [JsonPropertyName("workDoneItems")] public List<WorklogWorkDoneRecord> WorkDoneItems { get; set; } = new();
        [JsonPropertyName("photos")] public List<WorklogAttachmentRecord> Photos { get; set; } = new();
        [JsonPropertyName("files")] public List<WorklogAttachmentRecord> Files { get; set; } = new();
    }

    // ###########################################################################################
    // One row in an entry's "Links of interest" list: a free-text headline plus the URL it points
    // to.
    // ###########################################################################################
    public sealed class WorklogLinkRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("headline")] public string Headline { get; set; } = string.Empty;
        [JsonPropertyName("url")] public string Url { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // One row in an entry's "Comments" list. Every comment is user-added; the entry itself starts
    // with none.
    // ###########################################################################################
    public sealed class WorklogCommentRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("text")] public string Text { get; set; } = string.Empty;
        [JsonPropertyName("date")] public DateTime Date { get; set; }
    }

    // ###########################################################################################
    // One row in an entry's "Work done" list: a dated note plus the time spent and cost, so the
    // editor can show a running total across every row.
    // ###########################################################################################
    public sealed class WorklogWorkDoneRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("text")] public string Text { get; set; } = string.Empty;
        [JsonPropertyName("date")] public DateTime Date { get; set; }
        [JsonPropertyName("hoursSpent")] public double HoursSpent { get; set; }
        [JsonPropertyName("cost")] public double Cost { get; set; }

        // ###########################################################################################
        // The ISO 4217 code this row's Cost was ENTERED in, stamped when the row is written.
        //
        // WHY PER ROW rather than reading the current setting wherever a cost is shown: the currency
        // is one app-wide user preference that can be changed at any time, and Cost is a bare number.
        // Read the setting at display time and every historical figure is silently relabelled by that
        // change - a repair costed at 430 DKK starts reading "430 GBP" on screen and, far worse, in a
        // re-exported PDF that goes to someone who cannot ask. The number never moved; only the word
        // beside it did, and nothing recorded what it used to say.
        //
        // Blank means "written before this field existed", which every row on disk today reads as.
        // Those are displayed in the CURRENT setting exactly as they were before - there is no way
        // to know better, and inventing a code for them would be a guess presented as a record. See
        // WorklogCurrency.ResolveRecordedCode, which is the one place that fallback is decided.
        [JsonPropertyName("currencyCode")] public string CurrencyCode { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // One row in an entry's "Photos/images" or "Files" list: the attached file's own name (as
    // stored in the entry's "worklog_<id>" subfolder) plus a user comment. Photos and Files
    // use the same shape - only which list they sit in tells them apart. DisplayOrder lets photos
    // be dragged into a different order without renaming the files on disk.
    // ###########################################################################################
    public sealed class WorklogAttachmentRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("fileName")] public string FileName { get; set; } = string.Empty;
        [JsonPropertyName("comment")] public string Comment { get; set; } = string.Empty;
        [JsonPropertyName("displayOrder")] public int DisplayOrder { get; set; }
    }

    // ###########################################################################################
    // Reads and writes workbooks under the local "Workbooks" folder. EVERY RECORD IS ONE FOLDER
    // HOLDING ONE index.json, at both levels:
    //
    //   Workbooks/
    //   |-- index.json                  the root's own bookkeeping (the id counter below)
    //   |-- workbook_1/
    //   |   |-- index.json              the workbook
    //   |   |-- worklog_1/
    //   |   |   |-- index.json          worklog #1, all of it
    //   |   |   \-- 5v-rail.png         and its photos/files, beside it
    //   |   \-- worklog_2/
    //   \-- workbook_2/
    //
    // There is deliberately no file listing workbooks and none listing worklogs: every query scans
    // the folders on disk, so there is no bookkeeping file to keep in sync or go stale. DELETING A
    // FOLDER IS HOW EITHER IS REMOVED - including by the user, by hand, in a file manager: the UI
    // simply stops showing it, with no surviving row in a shared list to clean up afterwards. That
    // is why a worklog's record sits in the SAME folder as its attachments rather than in a list
    // beside them; a worklog is one folder, entirely.
    //
    // Worklogs used to live together in one "entries.json" per workbook, and workbook folders used
    // to be named by the bare id ("1"). Neither is read any more: there is deliberately NO migration
    // and NO fallback, so data written by an older build is not seen at all.
    //
    // Workbooks and single worklog entries can both be DELETED from the Workbooks tab (see
    // DeleteWorkbook and DeleteEntry), and neither one's id is ever handed out again.
    //
    // Ids are therefore NOT derived from what is on disk. Two persisted counters record the highest
    // id ever handed out - the root's own index.json for workbook ids, and each workbook's own
    // index.json (WorkbookRecord.LastWorklogId) for the entries inside it - so deleting #2 of two
    // workbooks leaves the next one as #3, with a gap where #2 was. Deriving the next id from the
    // surviving folders instead, which is what this used to do, re-handed a deleted number to the
    // next record: that silently makes an already-exported PDF ("Workbook_2_...") describe a
    // different repair, and it lets the new record inherit the deleted one's folder on disk,
    // attachments included. See the id-counter header further down for the full reasoning, and for
    // why there is deliberately no migration of workbooks written before the counters existed.
    //
    // Purely local: this folder sits beside the settings/log files, never inside the synced "Data"
    // folder - unless "--workbooks-root=" points it elsewhere, the same idea as DataManager's
    // "--data-root=" for the synced data. Call Load() once at startup before any other member is
    // used.
    // ###########################################################################################
    public static class WorklogManager
    {
        private const string WorkbooksRootArg = "--workbooks-root=";

        private static string _workbookRootPath = string.Empty;

        // The folder workbooks are actually being read from and written to - the "--workbooks-root="
        // value when one was given, otherwise the AppData default. The mirror of
        // DataManager.DataRoot, and exposed for the same reason: the Configuration tab's "Open
        // workbooks folder" button has to open the folder the app is really using, not the default
        // it would have used. Empty when Load has not run or failed.
        public static string WorkbookRoot => _workbookRootPath;

        // ###########################################################################################
        // Resolves the folder to store workbooks in - "--workbooks-root=" if given, otherwise the
        // "Workbooks" folder in the user's AppData folder - and points the manager at it.
        //
        // Falls back to an unusable (empty) root silently on any failure.
        // ###########################################################################################
        public static void Load(string[]? args = null)
        {
            try
            {
                string? explicitRoot = ResolveExplicitWorkbookRoot(args);

                if (explicitRoot != null)
                {
                    LoadFrom(explicitRoot);
                    return;
                }

                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var appFolder = Path.Combine(appData, AppConfig.AppFolderName);
                var directory = Path.Combine(appFolder, AppConfig.WorklogFolderName);

                LoadFrom(directory);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load worklog: [{ex.Message}] - using defaults");
            }
        }

        // ###########################################################################################
        // Parses "--workbooks-root=" out of the command line, following the same conventions
        // DataManager.ResolveDataRoot uses for "--data-root=": case-insensitive, surrounding quotes
        // stripped, first match wins. Returns null when not given, so Load() can fall back to its
        // AppData default.
        // ###########################################################################################
        internal static string? ResolveExplicitWorkbookRoot(string[]? args)
        {
            if (args == null)
                return null;

            foreach (var arg in args)
            {
                if (arg.StartsWith(WorkbooksRootArg, StringComparison.OrdinalIgnoreCase))
                    return arg[WorkbooksRootArg.Length..].Trim('"', '\'');
            }

            return null;
        }

        // ###########################################################################################
        // Points the manager at an explicit workbook root folder, creating it if missing. Load()
        // resolves the real AppData location and calls this; splitting the two lets the test suite
        // point at a temporary folder instead of the user's real one.
        // Falls back to an unusable (empty) root silently on any failure.
        // ###########################################################################################
        internal static void LoadFrom(string workbookRootPath)
        {
            try
            {
                _workbookRootPath = workbookRootPath;
                Directory.CreateDirectory(_workbookRootPath);

                Logger.Info($"Worklog loaded: [{ReadAllWorkbooks().Count} workbooks] from [{_workbookRootPath}]");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load worklog: [{ex.Message}] - using defaults");
                _workbookRootPath = string.Empty;
            }
        }

        // ###########################################################################################
        // Reads the id out of a "workbook_{id}" folder name, rejecting anything else.
        //
        // The rejection is what stops a worklog folder, an export, or any other stray directory in
        // the Workbooks root being read as a workbook. It also means a folder from before the rename
        // (a bare "1") is not seen at all, which is the no-fallback behaviour that was asked for.
        // ###########################################################################################
        private static bool TryParseWorkbookFolderName(string? folderName, out int workbookId)
        {
            workbookId = 0;

            if (folderName == null || !folderName.StartsWith(WorkbookFolderPrefix, StringComparison.Ordinal))
            {
                return false;
            }

            return int.TryParse(
                folderName[WorkbookFolderPrefix.Length..],
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out workbookId);
        }

        private const string WorkbookFolderPrefix = "workbook_";

        // ###########################################################################################
        // Reads every workbook subfolder's index.json under the workbook root. A subfolder that is
        // not named "workbook_{id}", or has no (or an unreadable) index.json, is skipped rather than
        // failing the whole read - it may be a workbook that is mid-delete, or debris left by hand.
        // ###########################################################################################
        private static List<WorkbookRecord> ReadAllWorkbooks()
        {
            var results = new List<WorkbookRecord>();

            if (string.IsNullOrEmpty(_workbookRootPath) || !Directory.Exists(_workbookRootPath))
            {
                return results;
            }

            foreach (var folder in Directory.GetDirectories(_workbookRootPath))
            {
                if (!TryParseWorkbookFolderName(Path.GetFileName(folder), out _))
                {
                    continue;
                }

                string indexPath = Path.Combine(folder, AppConfig.WorklogIndexFileName);
                if (!File.Exists(indexPath))
                {
                    continue;
                }

                try
                {
                    var json = File.ReadAllText(indexPath);
                    var record = JsonSerializer.Deserialize<WorkbookRecord>(json);
                    if (record != null)
                    {
                        results.Add(record);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to load workbook: [{folder}] [{ex.Message}] - skipped");
                }
            }

            return results;
        }

        // ###########################################################################################
        // Returns the id that CreateWorkbook will hand out next, for display before the workbook
        // actually exists (the "#1" preview in the create dialog). A PEEK, not a reservation: it does
        // not consume the number, so two dialogs open at once both preview the same one and whichever
        // saves first takes it - CreateWorkbook allocates for real at write time.
        //
        // The counter when there is one, the highest id on disk when there is not (an unreadable
        // file, or an install that predates the counter). Then skipped past anything already taken,
        // per the header above.
        //
        // AllocateNextWorkbookId calls this and then records the result, which is the only
        // difference between a peek and an allocation.
        // ###########################################################################################
        public static int PeekNextId() =>
            SkipIdsAlreadyOnDisk((ReadCounters()?.LastWorkbookId ?? HighestWorkbookIdOnDisk()) + 1);

        // ###########################################################################################
        // THE ID COUNTERS - why ids are never reused, and why that is worth a file of its own.
        //
        // Both id spaces (workbooks under the root, entries within one workbook) used to be
        // allocated as "highest currently on disk, plus one". That reuses the number of anything
        // deleted from the top: create #1 and #2, delete #2, and the next workbook is #2 again.
        //
        // That is not cosmetic. A workbook id names a PDF that has already been exported and very
        // likely emailed to a customer ("Workbook_2_Commodore_C64_20260904"), and both ids name real
        // folders on disk ("2/", "worklog_2/") that a recreated record would inherit the contents
        // of. So a reused number silently makes an old document describe a different repair.
        //
        // The fix is a counter that records the highest id ever HANDED OUT, rather than deriving it
        // from what survives. A deleted number is spent: delete #2 of two workbooks and the next is
        // #3, leaving a gap that correctly records that #2 existed. The workbook counter lives in
        // the root's own index.json; the per-workbook ENTRY counter lives in that workbook's own
        // index.json (WorkbookRecord.LastWorklogId), so it travels with the folder it numbers and is
        // deleted along with it - a workbook's entry ids only have to be unique within it.
        //
        // THERE IS DELIBERATELY NO MIGRATION of data written before these counters existed. Nothing
        // reads an old workbook to seed a counter and nothing rewrites one. What the allocators do
        // instead is REFUSE an id that is already taken on disk (see SkipIdsAlreadyOnDisk, and
        // AllocateNextEntryId for the entry side): a counter starting from zero would otherwise hand
        // out #1 while a "1" folder still sits there, and since CreateWorkbook's
        // Directory.CreateDirectory succeeds silently on an existing folder, the new workbook's
        // index.json would overwrite the old one in place and inherit its entries and attachments.
        // The skip costs one directory check per create and makes that unrepresentable. It is a
        // floor, not a migration - the counter is still the thing that stops a DELETED id coming
        // back, which is what checking disk alone can never do.
        // ###########################################################################################
        // The root's own index.json - NOT a workbook's. It sits directly in the workbook root while
        // every workbook's index.json sits one level down in its own numbered subfolder, so
        // ReadAllWorkbooks (which only ever walks subdirectories) never sees this one and cannot
        // mistake it for a workbook.
        //
        // Named index.json rather than the counters.json it used to be because "counters" plural
        // described a file holding one counter, and because this is the natural home for anything
        // else the root ever needs to record - the same role index.json already plays inside a
        // workbook folder. THERE IS DELIBERATELY NO MIGRATION: an install carrying the old
        // counters.json simply finds no index.json, falls back to the highest id on disk exactly as
        // a pre-counter install does, and writes the new file on its next allocation. That fallback
        // cannot collide with anything that exists (see SkipIdsAlreadyOnDisk); the only cost is that
        // ids deleted before the rename can come back once, which is the same exposure every
        // install had before the counter existed.
        private const string RootIndexFileName = AppConfig.WorklogIndexFileName;

        // ###########################################################################################
        // The workbook root's persisted id counter. One field today; a class rather than a bare int
        // in the JSON so a future counter can be added without rewriting the file's shape.
        // ###########################################################################################
        private sealed class WorkbookCounters
        {
            [JsonPropertyName("lastWorkbookId")] public int LastWorkbookId { get; set; }
        }

        private static string CountersPath =>
            string.IsNullOrEmpty(_workbookRootPath) ? string.Empty : Path.Combine(_workbookRootPath, RootIndexFileName);

        // ###########################################################################################
        // Reads the counter file, or null when it is missing or unreadable.
        //
        // An unreadable file is NOT treated as zero: that would restart numbering low and re-hand
        // ids that may already have been used. Callers fall back to the highest id on disk, which is
        // the old pre-counter behaviour - worse than the counter, but never a collision.
        // ###########################################################################################
        private static WorkbookCounters? ReadCounters()
        {
            string path = CountersPath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return null;
            }

            // A workbook's own record deserializes cleanly here - it simply has no lastWorkbookId,
            // so it would read as counter zero and restart numbering low. Same file, same name; see
            // IsWorkbookRecordFile for how the two come to share a path at all.
            if (IsWorkbookRecordFile(path))
            {
                return null;
            }

            try
            {
                return JsonSerializer.Deserialize<WorkbookCounters>(File.ReadAllText(path));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to read worklog id counters: [{path}] [{ex.Message}] - falling back to the highest id on disk");
                return null;
            }
        }

        // ###########################################################################################
        // Walks an id forward past every workbook folder that already exists, so a counter starting
        // from zero beside pre-existing workbooks can never hand out one of their numbers - see the
        // header above for what overwriting one would cost.
        //
        // A plain loop rather than "highest on disk plus one": the gaps left by deleted workbooks
        // are exactly what this must be able to walk INTO once the counter is ahead of them, so it
        // has to test the specific id rather than jump past everything.
        // ###########################################################################################
        private static int SkipIdsAlreadyOnDisk(int candidateId)
        {
            if (string.IsNullOrEmpty(_workbookRootPath) || !Directory.Exists(_workbookRootPath))
            {
                return candidateId;
            }

            while (Directory.Exists(Path.Combine(_workbookRootPath, BuildWorkbookFolderName(candidateId))))
            {
                candidateId++;
            }

            return candidateId;
        }

        // ###########################################################################################
        // Allocates the next workbook id and RECORDS it, so the number is spent whether or not the
        // workbook it was handed to still exists later.
        //
        // A failed counter write is logged and tolerated rather than failing the create: the user
        // came here to start a repair job, and refusing that over a bookkeeping file would be a
        // worse outcome than the id reuse it guards against. The next allocation then falls back to
        // max-on-disk, which still cannot collide with anything that exists.
        // ###########################################################################################
        private static int AllocateNextWorkbookId()
        {
            int id = PeekNextId();

            string path = CountersPath;
            if (!string.IsNullOrEmpty(path) && !IsWorkbookRecordFile(path))
            {
                var counters = ReadCounters() ?? new WorkbookCounters();
                counters.LastWorkbookId = id;
                AtomicJsonFile.Write(path, counters, SaveJsonOptions, "worklog id counters");
            }

            return id;
        }

        // ###########################################################################################
        // Whether the file at the root's index.json path is actually a WORKBOOK's own record rather
        // than this root's counter file.
        //
        // WHY THIS EXISTS: the counter file shares its name with a workbook record (both index.json,
        // see RootIndexFileName), and the only thing normally keeping them apart is which folder
        // they sit in - the counter at the root, records one level down. Point the workbook root AT
        // a workbook folder, though, and the two paths become the same file: --workbooks-root set to
        // a "workbook_3" folder by mistake, or any restore or hand-copy that nests the root one level
        // too deep. The first workbook created there would then write a counters object straight over
        // workbook #3's record, destroying its title, note, status, dates, schematic order and entry
        // counter in one write, with all of its worklog_N folders still on disk and now orphaned.
        //
        // A workbook record always carries a boardKey; the counter file has only lastWorkbookId, so
        // the presence of that property is the discriminator. An unreadable or non-JSON file is NOT
        // treated as a record - the counter has to be writable in the ordinary case, and a corrupt
        // counter file is exactly what the write is there to replace.
        //
        // Refusing the write rather than failing the create, for the reason AllocateNextWorkbookId
        // already gives: the fallback is max-id-on-disk, which is worse than a counter but never
        // collides with anything that exists. Losing "a deleted id stays spent" on a misconfigured
        // root is a far smaller cost than eating a workbook.
        // ###########################################################################################
        private static bool IsWorkbookRecordFile(string path)
        {
            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
                using var document = JsonDocument.Parse(File.ReadAllText(path));

                if (document.RootElement.ValueKind != JsonValueKind.Object)
                {
                    return false;
                }

                bool looksLikeARecord = document.RootElement.TryGetProperty("boardKey", out _);

                if (looksLikeARecord)
                {
                    Logger.Warning($"The worklog root's [{RootIndexFileName}] is a workbook record, not an id counter: [{path}] - the workbook root is very likely pointed at a workbook folder. The counter will not be written, and ids fall back to the highest on disk.");
                }

                return looksLikeARecord;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not inspect [{path}] before writing the worklog id counters: [{ex.Message}] - treating it as a counter file");
                return false;
            }
        }

        // ###########################################################################################
        // The highest numbered subfolder under the workbook root, or 0 when there are none. The
        // pre-counter allocation rule, kept only as the fallback described above.
        // ###########################################################################################
        private static int HighestWorkbookIdOnDisk()
        {
            if (string.IsNullOrEmpty(_workbookRootPath) || !Directory.Exists(_workbookRootPath))
            {
                return 0;
            }

            int maxId = 0;

            foreach (var folder in Directory.GetDirectories(_workbookRootPath))
            {
                if (TryParseWorkbookFolderName(Path.GetFileName(folder), out int id) && id > maxId)
                {
                    maxId = id;
                }
            }

            return maxId;
        }

        // ###########################################################################################
        // Returns the id AddEntry will hand out next for the given workbook, for display before the
        // entry actually exists: the on-board "#N" badge over a freshly drawn area, and the id a
        // draft entry in the full editor RESERVES for its attachment folder.
        //
        // Reads the workbook's persisted LastWorklogId rather than the entries currently in it, so a
        // deleted entry's number is never offered again - see the id-counter header above. A
        // workbook with no folder previews "#1".
        //
        // A PEEK, not a reservation: another entry can be added between this call and the save, so
        // AddEntryRecord re-allocates the id at write time rather than trusting the peeked one.
        // ###########################################################################################
        public static int PeekNextEntryId(int workbookId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                return 1;
            }

            return PeekNextEntryIdIn(folder);
        }

        // ###########################################################################################
        // The next entry id for an already-resolved workbook folder, WITHOUT consuming it.
        //
        // LastWorklogId is 0 both for a workbook that has never had an entry and for one written
        // before that field existed - and the second of those may hold entries numbered well above
        // zero. There is deliberately NO migration (nothing rewrites such a workbook to seed its
        // counter), so the id is instead walked past every entry actually present, exactly as the
        // workbook side skips existing folders. That is a floor, not a seed: once the counter is
        // ahead of what is on disk it does all the work, and a gap left by a deleted entry is
        // walked straight over rather than reused.
        //
        // The FOLDER is checked as well as the worklogs themselves. Every saved worklog now has a
        // folder, so for those two the checks agree; what the folder check still catches on its own
        // is a "worklog_{id}" folder holding no readable index.json - a cancelled draft whose
        // cleanup delete failed, or a worklog whose file was damaged. Handing that id out would
        // silently adopt a stranger's photos and files, and (since the record is written into that
        // same folder) drop the new worklog in on top of them.
        //
        // reservedId is the id a DRAFT already claimed and wrote its attachment bytes under, or 0
        // when there is no draft. It is EXEMPT from that folder check: that folder is not a
        // stranger's, it belongs to the worklog being saved right now, and skipping past it would
        // both misnumber the worklog and strand the photos the user just attached to it.
        // ###########################################################################################
        private static int PeekNextEntryIdIn(string folder, int reservedId = 0)
        {
            var record = ReadWorkbookRecord(folder);
            int candidateId = (record?.LastWorklogId ?? 0) + 1;

            var usedIds = new HashSet<int>(ReadEntries(folder).Select(e => e.Id));

            while (usedIds.Contains(candidateId) || IsWorklogFolderTaken(folder, candidateId, reservedId))
            {
                candidateId++;
            }

            return candidateId;
        }

        private static bool IsWorklogFolderTaken(string folder, int candidateId, int reservedId) =>
            candidateId != reservedId &&
            Directory.Exists(Path.Combine(folder, BuildEntryAttachmentsFolderName(candidateId)));

        // ###########################################################################################
        // Allocates the next entry id for a workbook and RECORDS it in that workbook's index.json,
        // so the number is spent whether or not the entry it was handed to still exists later.
        //
        // This writes index.json, and RecomputeWorkbookStatus writes it again at the end of the same
        // save. The two do not fight only because that one RE-READS the file rather than saving a
        // record captured earlier - see its own comment. Anything added later that rewrites this
        // file mid-save must do the same, or it will silently roll the counter back and re-hand the
        // id that was just allocated.
        //
        // A failed write is logged and tolerated rather than failing the save: the user's typed work
        // matters more than the bookkeeping field, and the walk-past-what-exists floor in
        // PeekNextEntryIdIn still prevents a collision with anything that is actually there.
        // ###########################################################################################
        private static int AllocateNextEntryId(string folder, int workbookId, int reservedId = 0)
        {
            int id = PeekNextEntryIdIn(folder, reservedId);

            var record = ReadWorkbookRecord(folder);
            if (record != null)
            {
                record.LastWorklogId = id;
                SaveWorkbook(folder, record);
            }
            else
            {
                Logger.Warning($"Could not record the allocated worklog id [#{id}] for workbook [#{workbookId}]: its index.json could not be read");
            }

            return id;
        }

        // ###########################################################################################
        // One workbook's own index.json, or null when it is missing or unreadable. The single reader
        // for that file - UpdateWorkbook, RecomputeWorkbookStatus and the entry-id counter each
        // used to open and deserialize it inline with their own try/catch.
        // ###########################################################################################
        private static WorkbookRecord? ReadWorkbookRecord(string folder)
        {
            string indexPath = Path.Combine(folder, AppConfig.WorklogIndexFileName);
            if (!File.Exists(indexPath))
            {
                return null;
            }

            try
            {
                return JsonSerializer.Deserialize<WorkbookRecord>(File.ReadAllText(indexPath));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to read workbook: [{folder}] [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // Returns EVERY workbook recorded for the given board key, newest first, or an empty list
        // when the board has none.
        //
        // The two lookups below each collapse a board's history to a single workbook - the active
        // one, or the latest one - because the worklog bar has room for exactly one. The Workbooks
        // tab lists the whole history instead, so it needs the unreduced set. Ordering matches
        // theirs (descending id, so newest first) rather than by date: the id is what the user sees
        // on each card as "#12", and a list numbered 12, 11, 9 that was not in that order would
        // look sorted by nothing at all.
        //
        // An empty or unknown board key yields an empty list rather than every workbook on disk.
        // Returning everything would show one board's repairs under another board's name, which is
        // worse than showing nothing.
        // ###########################################################################################
        public static List<WorkbookRecord> GetWorkbooksForBoard(string boardKey)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
            {
                return new List<WorkbookRecord>();
            }

            return ReadAllWorkbooks()
                .Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal))
                .OrderByDescending(w => w.Id)
                .ToList();
        }

        // ###########################################################################################
        // Returns EVERY workbook on disk, for EVERY board, newest first - the one place a caller
        // genuinely wants workbooks that are not scoped to "the current board" (the worklog bar's
        // picker, so a workbook for a different board can be selected without a trip to that board
        // first). Everywhere else in the app - the Workbooks tab, GetActiveWorkbookForBoard,
        // ResolveActiveWorkbook - deliberately stays board-scoped; see GetWorkbooksForBoard's own
        // header for why an unscoped list is usually the wrong answer.
        // ###########################################################################################
        public static List<WorkbookRecord> GetAllWorkbooks() =>
            ReadAllWorkbooks().OrderByDescending(w => w.Id).ToList();

        // ###########################################################################################
        // Returns the most recently created still-open workbook for the given board key, or null
        // when that board has no open workbook (including when it has none at all, or its folder
        // was deleted). A board can accumulate several closed workbooks over time; only the
        // highest-id open one counts as "active".
        //
        // NOT the lookup any UI should use, despite the name. "Which workbook is active" is
        // ResolveActiveWorkbook above: it honours the workbook the user activated on the Workbooks
        // tab, and it is status-blind, so a finished workbook stays visible instead of vanishing
        // from the UI the moment its last entry is resolved. Nothing in the app calls this any
        // more - it is kept as a query the WorklogManager tests use to assert the OPEN/CLOSED
        // bookkeeping RecomputeWorkbookStatus does. Calling it from a new worklog feature would
        // silently reintroduce "newest open one, ignore what the user activated".
        // ###########################################################################################
        public static WorkbookRecord? GetActiveWorkbookForBoard(string boardKey)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
            {
                return null;
            }

            return ReadAllWorkbooks()
                .Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal) &&
                            string.Equals(w.Status, "Open", StringComparison.Ordinal))
                .OrderByDescending(w => w.Id)
                .FirstOrDefault();
        }

        // ###########################################################################################
        // Returns the most recently created workbook for the given board key regardless of its
        // status, or null when that board has none at all.
        //
        // Status-blind, deliberately unlike GetActiveWorkbookForBoard above: resolving a workbook's
        // last outstanding entry closes it, and when the bar only ever asked for an OPEN workbook
        // that made the whole workbook vanish from the UI the moment it was finished - it looked
        // like data loss even though everything was still on disk.
        //
        // NOT the lookup any UI should use either, and for the opposite reason to the one above:
        // "newest wins, always" is exactly what activating an older or closed workbook on the
        // Workbooks tab is meant to override. Every former caller (the worklog bar, "Show worklogs",
        // "Add worklog") now goes through ResolveActiveWorkbook. Kept only as the query the
        // WorklogManager tests use to read a board's newest workbook back off disk; calling it from
        // a new worklog feature would silently ignore what the user activated.
        // ###########################################################################################
        public static WorkbookRecord? GetLatestWorkbookForBoard(string boardKey)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
            {
                return null;
            }

            return ReadAllWorkbooks()
                .Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal))
                .OrderByDescending(w => w.Id)
                .FirstOrDefault();
        }

        // ###########################################################################################
        // Picks the ONE workbook every worklog-facing surface acts on for a board: the worklog bar,
        // "Show worklogs", "Add worklog", and the Workbooks tab's highlighted card and board pane.
        //
        // The saved id (UserSettings.ActiveWorkbookIdByBoard, set when the user clicks a card in the
        // Workbooks tab) wins whenever it still names a workbook this board actually has; otherwise
        // the board's newest, which is what everything defaulted to before workbooks could be
        // activated at all. The saved id is validated on every call rather than trusted: a workbook
        // folder can be deleted by hand, and an unvalidated id would leave the bar quietly showing
        // nothing instead of falling back.
        //
        // Pure, and takes its two inputs rather than reading them: Main and TabWorkbooks each had
        // their own copy of this rule with different shapes (one returning the record, one the id),
        // which is precisely the disagreement between the highlighted card and the bar that the
        // design was supposed to prevent. One implementation, unit-testable, called from both.
        //
        // "workbooks" is expected newest-first, which is GetWorkbooksForBoard's own order.
        // ###########################################################################################
        public static WorkbookRecord? ResolveActiveWorkbook(IReadOnlyList<WorkbookRecord> workbooks, int? savedActiveId)
        {
            if (workbooks == null || workbooks.Count == 0)
            {
                return null;
            }

            if (savedActiveId.HasValue)
            {
                var saved = workbooks.FirstOrDefault(w => w.Id == savedActiveId.Value);
                if (saved != null)
                {
                    return saved;
                }
            }

            return workbooks[0];
        }

        // ###########################################################################################
        // Whether an entry state counts as finished. The ONE place that question is answered, so a
        // second resolved state added to ResolvedEntryStates cannot auto-close a workbook while
        // pills, padlocks and colours elsewhere still draw its entries as open.
        //
        // Case-insensitive, unlike the set it consults: the set decides the auto-close rule against
        // states this app wrote itself, while this is asked about states read back off disk, which
        // can carry "closed" from a hand edit or an older build. Falling through to "open" there
        // renders a red padlock on a pill whose own label reads "closed", and is indistinguishable
        // from the intended default - so it would never be noticed.
        // ###########################################################################################
        public static bool IsResolvedState(string? state) =>
            !string.IsNullOrWhiteSpace(state) &&
            ResolvedEntryStates.Any(resolved => string.Equals(resolved, state.Trim(), StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // Whether a WORKBOOK's status counts as open - the other axis, the one RecomputeWorkbookStatus
        // writes. Separate from IsResolvedState above because they answer different questions about
        // different fields, even though both read "Open"/"Closed" to the user.
        //
        // Anything that is not recognisably "Closed" reads as open, matching how every status pill in
        // the app already falls back, and case-insensitively for the same reason IsResolvedState is:
        // a status read back off disk can carry other casing.
        // ###########################################################################################
        public static bool IsWorkbookStatusOpen(string? status) =>
            !string.Equals(status?.Trim(), "Closed", StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Creates a new open workbook for the given board: allocates the next id, creates its own
        // subfolder under the workbook root, and writes that workbook's index.json into it.
        // Returns null when the workbook root is unusable or the folder cannot be written.
        //
        // The empty-root guard matters as much here as in the readers: LoadFrom blanks the root on
        // any failure, and Path.Combine("", "1") yields the *relative* path "1", so without it this
        // silently created workbook folders next to the executable. Those looked fine in the bar but
        // were invisible to GetWorkbookFolder, which does guard - so every entry the user then
        // recorded was discarded with nothing but a log line.
        // ###########################################################################################
        public static WorkbookRecord? CreateWorkbook(string boardKey, string title, string note)
        {
            if (string.IsNullOrEmpty(_workbookRootPath))
            {
                Logger.Warning("Failed to create workbook: no usable workbook root folder");
                return null;
            }

            // Allocates AND records, so this number is spent even if the workbook is later
            // deleted - see the id-counter header above.
            int id = AllocateNextWorkbookId();
            string folder = Path.Combine(_workbookRootPath, BuildWorkbookFolderName(id));

            var record = new WorkbookRecord
            {
                Id = id,
                BoardKey = boardKey,
                Title = title.Trim(),
                Note = note?.Trim() ?? string.Empty,
                Status = "Open",
                StartDate = DateTime.Now.Date,
                WorklogCount = 0
            };

            try
            {
                Directory.CreateDirectory(folder);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to create workbook folder: [{folder}] [{ex.Message}]");
                return null;
            }

            if (!SaveWorkbook(folder, record))
            {
                return null;
            }

            Logger.Info($"Setting changed: [Worklog] created workbook [#{id}] [{record.Title}] for board [{boardKey}]");

            return record;
        }

        // ###########################################################################################
        // Overwrites an existing workbook's title/note in place (its id, board key, status, start
        // date and entry count are untouched - this only edits the two fields the create dialog
        // collects, reused for "Edit workbook" via CreateWorkbookWindow.InitializeForEdit).
        //
        // Returns the SAVED record, or null when the workbook cannot be found or the write itself
        // failed - the same shape CreateWorkbook returns, and for the same reason: the caller needs
        // the record that actually reached disk. Returning a bool instead left the caller patching up
        // its own in-memory copy field by field, re-applying the trim rules a second time, so the two
        // could drift apart the moment either side's trimming changed.
        // ###########################################################################################
        public static WorkbookRecord? UpdateWorkbook(int workbookId, string title, string note)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to update workbook: [#{workbookId}] folder not found");
                return null;
            }

            var record = ReadWorkbookRecord(folder);
            if (record == null)
            {
                Logger.Warning($"Failed to update workbook [#{workbookId}]: its index.json could not be read");
                return null;
            }

            record.Title = title?.Trim() ?? string.Empty;
            record.Note = note?.Trim() ?? string.Empty;

            if (!SaveWorkbook(folder, record))
            {
                return null;
            }

            Logger.Info($"Setting changed: [Worklog] updated workbook [#{workbookId}] [{record.Title}]");

            return record;
        }

        // ###########################################################################################
        // Saves the workbook's schematic preview order (see WorkbookRecord.SchematicOrder), leaving
        // every other field alone. Returns false when the workbook cannot be found or the write
        // failed.
        //
        // Its own method rather than a parameter on UpdateWorkbook: that one is the "Edit workbook"
        // dialog's save and takes the title and note the dialog collects, so threading an unrelated
        // third field through it would make every caller pass two values it has no opinion about.
        // This one re-reads the record first for the same reason UpdateWorkbook does - a caller's
        // copy can be stale, and writing it back would roll off whatever has been saved since.
        // ###########################################################################################
        public static bool UpdateWorkbookSchematicOrder(int workbookId, List<string> schematicOrder)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to save schematic order: workbook [#{workbookId}] folder not found");
                return false;
            }

            var record = ReadWorkbookRecord(folder);
            if (record == null)
            {
                Logger.Warning($"Failed to save schematic order for workbook [#{workbookId}]: its index.json could not be read");
                return false;
            }

            record.SchematicOrder = schematicOrder ?? new List<string>();

            return SaveWorkbook(folder, record);
        }

        // ###########################################################################################
        // Deletes a workbook entirely. Its index.json and every one of its worklogs - each its own
        // "worklog_<id>" folder holding that worklog's record and attachments - all live inside the
        // workbook's own folder (see the class header), so removing that one folder removes the
        // whole workbook, with no separate bookkeeping anywhere else that could be left dangling.
        // Returns false when the workbook cannot be found or the folder could not be removed (e.g. a
        // file inside it is locked open elsewhere).
        // ###########################################################################################
        public static bool DeleteWorkbook(int workbookId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to delete workbook: [#{workbookId}] folder not found");
                return false;
            }

            try
            {
                Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to delete workbook [#{workbookId}]: [{ex.Message}]");
                return false;
            }

            Logger.Info($"Setting changed: [Worklog] deleted workbook [#{workbookId}]");

            return true;
        }

        // ###########################################################################################
        // Resolution states that count as "closed" for a single entry. A workbook auto-closes once
        // every one of its entries is in one of these states - see RecomputeWorkbookStatus.
        //
        // An entry has just two states now, Open and Closed, matching the workbook's own vocabulary.
        // A set (rather than a plain equality check) is kept because the auto-close rule is about
        // "which states mean finished", which is the thing likeliest to grow again.
        // ###########################################################################################
        private static readonly HashSet<string> ResolvedEntryStates = new(StringComparer.Ordinal) { "Closed" };

        // ###########################################################################################
        // The automatic comments the app writes into an entry's own Comments list when something
        // worth recording happens to it: it was created, its state was flipped, or its category was
        // changed. They read as an audit trail beside the user's own comments, so a worklog explains
        // its own history rather than only showing its current state.
        //
        // The wording lives here rather than at the two call sites (the quick create card and the
        // full editor) so the phrasing cannot drift apart between them, and so the tests assert the
        // same strings the app writes.
        // ###########################################################################################
        public const string CreatedCommentText = "Worklog created";

        public const string OpenedCommentText = "Worklog opened";

        public const string ClosedCommentText = "Worklog closed";

        // The category is quoted, deliberately: a bare Worklog changed to Note reads as a sentence
        // that has lost a word, and the quotes mark the value as the literal category name.
        public static string BuildCategoryChangedCommentText(string category) =>
            $"Worklog changed to \"{category}\"";

        // The state comment for a state value - null for anything that is neither Open nor Closed,
        // so an unrecognised state (a value from a future build, say) records nothing rather than
        // claiming the worklog was opened.
        public static string? BuildStateChangedCommentText(string state)
        {
            if (string.Equals(state, "Open", StringComparison.Ordinal))
                return OpenedCommentText;

            if (string.Equals(state, "Closed", StringComparison.Ordinal))
                return ClosedCommentText;

            return null;
        }

        // ###########################################################################################
        // Appends an automatic comment to the given list, allocating the next free id the same way
        // the editor's own Add-comment does. Returns the record it added so a caller can show it.
        //
        // Blank text adds nothing: BuildStateChangedCommentText can decline to describe a state, and
        // an empty comment row in the list would be worse than no row at all.
        // ###########################################################################################
        public static WorklogCommentRecord? AppendAutomaticComment(List<WorklogCommentRecord> comments, string? text)
        {
            if (comments == null || string.IsNullOrWhiteSpace(text))
                return null;

            int nextId = comments.Count == 0 ? 1 : comments.Max(c => c.Id) + 1;

            var comment = new WorklogCommentRecord
            {
                Id = nextId,
                Text = text.Trim(),
                Date = DateTime.Now
            };

            comments.Add(comment);
            return comment;
        }

        // ###########################################################################################
        // Every worklog in the given workbook, read one "worklog_<id>/index.json" at a time (empty
        // list when the workbook has none yet, or its folder cannot be found).
        // ###########################################################################################
        public static List<WorklogEntryRecord> GetEntries(int workbookId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                return new List<WorklogEntryRecord>();
            }

            return ReadEntries(folder);
        }

        // ###########################################################################################
        // Reads every worklog in a workbook - one "worklog_{id}/index.json" per worklog, discovered
        // by scanning the workbook's folder.
        //
        // ONE FILE PER WORKLOG, INSIDE THE FOLDER THAT ALREADY HOLDS ITS PHOTOS AND FILES. There is
        // deliberately no list file: the user can delete a worklog by deleting its folder in the
        // file manager and the UI simply stops showing it, with no surviving row in a shared list to
        // clean up by hand. That was the whole point of the change, and it is why the worklog's data
        // sits in the SAME folder as its attachments rather than in one of its own - a worklog is
        // one folder, entirely.
        //
        // A folder that is not named "worklog_{id}", or holds no readable index.json, is skipped
        // rather than failing the whole read - the same tolerance ReadAllWorkbooks applies one level
        // up, and it covers a worklog mid-delete as well as any stray folder.
        //
        // The id comes from the FOLDER NAME, not from the record: the folder is what names the
        // worklog's attachments, so a record whose stored id disagreed (a hand-edited file, a folder
        // copied to a new name) would otherwise point the UI at photos belonging to a different
        // worklog. The folder wins, and the record is corrected to match.
        //
        // Sorted by id so callers get a stable, counting order - directory enumeration order is
        // whatever the filesystem returns, and the entries used to arrive in the order they were
        // appended to entries.json.
        //
        // Deliberately NO fallback to the old entries.json: a workbook written by an older build
        // reads as having no worklogs, which is the no-migration behaviour that was asked for.
        // ###########################################################################################
        private static List<WorklogEntryRecord> ReadEntries(string folder)
        {
            var entries = new List<WorklogEntryRecord>();

            if (!Directory.Exists(folder))
            {
                return entries;
            }

            foreach (var worklogFolder in Directory.GetDirectories(folder))
            {
                if (!TryParseEntryFolderName(Path.GetFileName(worklogFolder), out int entryId))
                {
                    continue;
                }

                var entry = ReadEntryRecord(worklogFolder);
                if (entry == null)
                {
                    continue;
                }

                entry.Id = entryId;
                entries.Add(entry);
            }

            entries.Sort((left, right) => left.Id.CompareTo(right.Id));
            return entries;
        }

        // ###########################################################################################
        // One worklog's own index.json, or null when it is missing or unreadable.
        // ###########################################################################################
        private static WorklogEntryRecord? ReadEntryRecord(string worklogFolder)
        {
            string indexPath = Path.Combine(worklogFolder, AppConfig.WorklogIndexFileName);
            if (!File.Exists(indexPath))
            {
                return null;
            }

            try
            {
                var entry = JsonSerializer.Deserialize<WorklogEntryRecord>(File.ReadAllText(indexPath));
                if (entry == null)
                {
                    return null;
                }

                NormalizeEntryCollections(entry);
                entry.State = MigrateEntryState(entry.State);
                return entry;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load worklog: [{worklogFolder}] [{ex.Message}] - skipped");
                return null;
            }
        }

        // ###########################################################################################
        // Reads the id out of a "worklog_{id}" folder name, rejecting anything else - the worklog-level
        // twin of TryParseWorkbookFolderName, and what keeps any other folder inside a workbook from
        // being read as a worklog.
        // ###########################################################################################
        private static bool TryParseEntryFolderName(string? folderName, out int entryId)
        {
            entryId = 0;

            if (folderName == null || !folderName.StartsWith(EntryFolderPrefix, StringComparison.Ordinal))
            {
                return false;
            }

            return int.TryParse(
                folderName[EntryFolderPrefix.Length..],
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out entryId);
        }

        private const string EntryFolderPrefix = "worklog_";

        // ###########################################################################################
        // Maps the retired entry states onto the two that replaced them.
        //
        // Entries used to carry Pending/RuledOut/Fixed; they now carry Open/Closed, matching the
        // vocabulary a workbook already used. Without this mapping every entry saved by an older
        // build would be silently misreported: "Fixed" and "RuledOut" both counted as resolved, so
        // dropping them from ResolvedEntryStates reopens workbooks the user had finished, the
        // editor renders NEITHER state pill as selected (it only knows the new two), and saving
        // writes the retired value straight back.
        //
        // Ruled out maps to Closed rather than Open on purpose: it counted as resolved under the
        // old rule, so mapping it anywhere else would change every affected workbook's status
        // rather than preserve it. The mapping runs on read, so it costs nothing once the value has
        // been written back by any later save, and an unrecognised value is left alone for
        // RecomputeWorkbookStatus to treat as unresolved (the safe direction - a workbook stays
        // open rather than silently closing).
        // ###########################################################################################
        internal static string MigrateEntryState(string? state)
        {
            string trimmed = state?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(trimmed))
            {
                return "Open";
            }

            if (string.Equals(trimmed, "Pending", StringComparison.OrdinalIgnoreCase))
            {
                return "Open";
            }

            if (string.Equals(trimmed, "Fixed", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "RuledOut", StringComparison.OrdinalIgnoreCase))
            {
                return "Closed";
            }

            return trimmed;
        }

        // ###########################################################################################
        // Replaces any null collection on a freshly deserialized entry with an empty one.
        //
        // System.Text.Json assigns straight over a property's "= new()" initializer when the JSON
        // carries an explicit null, so a hand-edited or partially-written worklog file yields an
        // entry whose ComponentLabels/Links/Comments/WorkDoneItems/Photos/Files are null despite
        // their initializers. Every consumer then dereferences them unguarded - the editor's
        // CloneEntry did so six times on the path that opens a saved entry. Fixing it here means no
        // caller has to know, and none can forget.
        // ###########################################################################################
        private static void NormalizeEntryCollections(WorklogEntryRecord entry)
        {
            entry.ComponentLabels ??= new List<string>();
            entry.Links ??= new List<WorklogLinkRecord>();
            entry.Comments ??= new List<WorklogCommentRecord>();
            entry.WorkDoneItems ??= new List<WorklogWorkDoneRecord>();
            entry.Photos ??= new List<WorklogAttachmentRecord>();
            entry.Files ??= new List<WorklogAttachmentRecord>();
        }

        // ###########################################################################################
        // The folder name one workbook's data lives in - "workbook_{id}".
        //
        // Named for what it holds, matching BuildEntryAttachmentsFolderName's "worklog_{id}" one
        // level down, so the Workbooks folder says what each folder is when a user opens it rather
        // than presenting a row of bare numbers. It was the bare id.
        //
        // Deliberately NO migration and NO fallback to the old name: a folder from an older build
        // keeps its numeric name and simply stops being seen, exactly as the worklog attachment
        // folders did when they were renamed.
        // ###########################################################################################
        public static string BuildWorkbookFolderName(int workbookId) =>
            $"workbook_{workbookId.ToString(CultureInfo.InvariantCulture)}";

        // ###########################################################################################
        // Finds the id'd workbook's own subfolder by its name (see the class header for why there is
        // no other lookup), returning null when it does not exist.
        // ###########################################################################################
        private static string? GetWorkbookFolder(int workbookId)
        {
            if (string.IsNullOrEmpty(_workbookRootPath))
            {
                return null;
            }

            string folder = Path.Combine(_workbookRootPath, BuildWorkbookFolderName(workbookId));
            return Directory.Exists(folder) ? folder : null;
        }

        // ###########################################################################################
        // Appends an ALREADY-BUILT entry to the given workbook - the shape AddEntry's own field list
        // cannot express, because the record arrives carrying its sub-lists (links, comments, work
        // done, photos, files) already populated.
        //
        // This is what the "Add worklog" flow writes. Drawing an area on the schematic now opens the
        // FULL editor directly rather than a small quick card, and that editor holds the whole entry
        // - sub-lists included - in memory until Save, so the record reaching disk for the first time
        // is not the bare title/description/category/state one AddEntry takes. See
        // WorklogEntryEditorWindow.InitializeForNewEntry.
        //
        // The id is ASSIGNED HERE rather than taken from the record. The editor needed an id up front
        // (its attachment folder is named after one - see GetEntryAttachmentsFolder), and it got that
        // from PeekNextEntryId; but a peek is not a reservation, so between the peek and this call
        // another entry may have been added to the same workbook. Re-allocating from the entries as
        // they are NOW is what stops two entries sharing an id, which entries.json has no way to
        // represent and UpdateEntry would then resolve to whichever came first.
        //
        // reservedId reports the id the caller had been using, so it can move any attachment bytes it
        // already wrote under that id into the folder for the id actually allocated. It is almost
        // always the same number; when it is not, the bytes would otherwise be stranded in a folder
        // no entry names.
        //
        // Returns the saved entry (carrying its final id), or null when the workbook folder cannot be
        // found or the write itself failed - in both cases nothing was persisted.
        // ###########################################################################################
        public static WorklogEntryRecord? AddEntryRecord(int workbookId, WorklogEntryRecord entry, int reservedId)
        {
            if (entry == null)
            {
                return null;
            }

            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to add worklog: workbook [#{workbookId}] folder not found");
                return null;
            }

            var entries = ReadEntries(folder);
            // Allocates AND records in the workbook's index.json, so this number is spent even if
            // the entry is later deleted - see the id-counter header above.
            int nextEntryId = AllocateNextEntryId(folder, workbookId, reservedId);

            entry.Id = nextEntryId;
            NormalizeEntryCollections(entry);

            // Trimmed exactly as AddEntry does. The editor's Save gate uses IsNullOrWhiteSpace, so a
            // title of "  CPU socket  " passes it; writing that verbatim would sort and search
            // differently from every entry AddEntry ever wrote, and render its padding on the card.
            entry.Title = entry.Title?.Trim() ?? string.Empty;
            entry.Description = entry.Description?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(entry.Category))
            {
                entry.Category = "Note";
            }

            if (string.IsNullOrWhiteSpace(entry.State))
            {
                entry.State = "Open";
            }

            if (entry.CreatedDate == default)
            {
                entry.CreatedDate = DateTime.Now;
            }

            // The attachment bytes the editor wrote while the entry was still a draft live in the
            // folder its RESERVED id names. Moved BEFORE the record is written, not after: the
            // record names those files, so committing it first and then failing to move them
            // publishes a worklog whose photo and file rows point at a folder that does not hold
            // them. Doing it first means a failure leaves the worklog unwritten instead - the user
            // still has everything on screen and can retry Save.
            //
            // It also has to happen first for a second reason now: the record is written INTO that
            // same folder (see SaveEntry), so moving afterwards would have to move a file that was
            // just written under the wrong id.
            //
            // A failure is logged and not fatal: the worklog itself still saves, since the
            // alternative would lose the user's typed work over attachments they can re-add.
            if (reservedId != nextEntryId)
            {
                MoveEntryAttachmentsFolder(folder, reservedId, nextEntryId);
            }

            if (!SaveEntry(folder, entry))
            {
                // Nothing reached disk - returning the entry would tell the caller it was saved.
                return null;
            }

            entries.Add(entry);
            RecomputeWorkbookStatus(folder, workbookId, entries);

            Logger.Info($"Setting changed: [Worklog] added worklog [#{nextEntryId}] to workbook [#{workbookId}] [{entry.Category}/{entry.State}]");

            return entry;
        }

        // ###########################################################################################
        // Moves a draft entry's attachment folder to the id the entry was actually saved under - see
        // AddEntryRecord for when the two differ. Does nothing when the draft folder was never
        // created (the overwhelmingly common case: no attachment was added).
        //
        // A destination that ALREADY EXISTS is merged into, file by file, rather than skipped. It
        // can exist perfectly legitimately - a previous draft that reserved this same number and
        // whose cleanup delete failed leaves one behind - and skipping stranded the draft's bytes in
        // a folder no entry names, while the entry itself was saved naming filenames that were not
        // in its own folder. The filenames cannot collide across the two: they are built from the
        // owning record's own attachment ids (see WorklogAttachmentStorage.BuildStoredFileName), and
        // AllocateAttachmentId already refuses an id whose file is sitting in the folder. A name
        // that collides anyway is left alone rather than overwritten - the file already there is the
        // one some record may still name.
        // ###########################################################################################
        private static void MoveEntryAttachmentsFolder(string workbookFolder, int fromEntryId, int toEntryId)
        {
            string from = Path.Combine(workbookFolder, BuildEntryAttachmentsFolderName(fromEntryId));
            string to = Path.Combine(workbookFolder, BuildEntryAttachmentsFolderName(toEntryId));

            try
            {
                if (!Directory.Exists(from))
                {
                    return;
                }

                Directory.CreateDirectory(to);

                // FILE BY FILE, NEVER Directory.Move OF THE WHOLE FOLDER. The source folder is the
                // one the draft reserved, and that number can meanwhile have been taken by a
                // worklog that is actually SAVED - which is the very reason this method is being
                // called. That worklog's own index.json now lives in this folder, so renaming the
                // folder wholesale would carry the saved worklog's record away with the draft's
                // photos and delete it from the workbook. It did exactly that, and
                // AddEntryRecord_reallocates_the_id_when_another_entry_claimed_the_reserved_one
                // is the test that caught it.
                foreach (string sourceFile in Directory.GetFiles(from))
                {
                    string name = Path.GetFileName(sourceFile);

                    // The record belongs to whichever worklog owns this folder, not to the draft's
                    // attachments - the draft's own record has not been written yet and is about to
                    // be saved into the destination by the caller.
                    if (string.Equals(name, AppConfig.WorklogIndexFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string targetFile = Path.Combine(to, name);

                    if (File.Exists(targetFile))
                    {
                        Logger.Warning(
                            $"Not moving worklog draft attachment [{name}]: a file of that name already exists in [{BuildEntryAttachmentsFolderName(toEntryId)}]");
                        continue;
                    }

                    File.Move(sourceFile, targetFile);
                }

                // Only when it is actually empty - anything left behind is a file that could not be
                // moved (or the index.json of a saved worklog that took this number), and deleting
                // the folder would delete it.
                if (Directory.GetFileSystemEntries(from).Length == 0)
                {
                    Directory.Delete(from);
                }

                Logger.Info($"Moved worklog draft attachments from [{BuildEntryAttachmentsFolderName(fromEntryId)}] into [{BuildEntryAttachmentsFolderName(toEntryId)}]");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to move worklog draft attachments from [{from}] to [{to}]: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Appends a new entry to the given workbook: allocates the next entry id (scoped to this
        // workbook's own worklogs), writes it as its own "worklog_<id>/index.json", then recomputes
        // and saves the workbook's own status/entryCount -
        // see RecomputeWorkbookStatus for the Open/Closed rule.
        // Returns the saved entry, or null when the workbook folder cannot be found or the write
        // itself failed (a full or locked disk) - in both cases nothing was persisted, so the
        // caller must not treat it as saved.
        // ###########################################################################################
        public static WorklogEntryRecord? AddEntry(
            int workbookId,
            string schematicName,
            Rect area,
            string title,
            string description,
            string category,
            string state,
            IEnumerable<string> componentLabels,
            bool showMarkedArea = true)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to add worklog: workbook [#{workbookId}] folder not found");
                return null;
            }

            var entries = ReadEntries(folder);
            // Allocates AND records in the workbook's index.json, so this number is spent even if
            // the entry is later deleted - see the id-counter header above.
            int nextEntryId = AllocateNextEntryId(folder, workbookId);

            var entry = new WorklogEntryRecord
            {
                Id = nextEntryId,
                SchematicName = schematicName ?? string.Empty,
                AreaX = area.X,
                AreaY = area.Y,
                AreaWidth = area.Width,
                AreaHeight = area.Height,
                Title = title?.Trim() ?? string.Empty,
                Description = description?.Trim() ?? string.Empty,
                Category = string.IsNullOrWhiteSpace(category) ? "Note" : category,
                State = string.IsNullOrWhiteSpace(state) ? "Open" : state,
                ComponentLabels = componentLabels?.ToList() ?? new List<string>(),
                ShowMarkedArea = showMarkedArea,
                CreatedDate = DateTime.Now
            };

            // Every worklog starts its own history with the fact that it was created, so the
            // Comments list is an audit trail from the first entry rather than starting empty and
            // only recording what happened later.
            AppendAutomaticComment(entry.Comments, CreatedCommentText);

            if (!SaveEntry(folder, entry))
            {
                // Nothing reached disk - returning the entry would tell the caller it was saved.
                return null;
            }

            entries.Add(entry);
            RecomputeWorkbookStatus(folder, workbookId, entries);

            Logger.Info($"Setting changed: [Worklog] added worklog [#{nextEntryId}] to workbook [#{workbookId}] [{entry.Category}/{entry.State}]");

            return entry;
        }

        // ###########################################################################################
        // Overwrites an existing entry in place (matched by Id) with everything the full editor can
        // change: title, description, category, state, and the Links/Comments/WorkDoneItems/Photos/
        // Files sub-lists. Then recomputes the workbook's status the same way AddEntry does, since
        // editing an entry's State is exactly how the user resolves (or reopens) it from the editor.
        // Returns false when the workbook or the entry itself cannot be found.
        // ###########################################################################################
        public static bool UpdateEntry(int workbookId, WorklogEntryRecord updatedEntry)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to update worklog: workbook [#{workbookId}] folder not found");
                return false;
            }

            var entries = ReadEntries(folder);
            int index = entries.FindIndex(e => e.Id == updatedEntry.Id);
            if (index < 0)
            {
                Logger.Warning($"Failed to update worklog: worklog [#{updatedEntry.Id}] not found in workbook [#{workbookId}]");
                return false;
            }

            if (!SaveEntry(folder, updatedEntry))
            {
                // Nothing reached disk - reporting true here is what made the editor close looking
                // saved while the user's edits reverted on the next refresh.
                return false;
            }

            entries[index] = updatedEntry;

            RecomputeWorkbookStatus(folder, workbookId, entries);

            Logger.Info($"Setting changed: [Worklog] updated worklog [#{updatedEntry.Id}] in workbook [#{workbookId}] [{updatedEntry.Category}/{updatedEntry.State}]");

            return true;
        }

        // ###########################################################################################
        // Deletes one worklog from a workbook by removing its whole "worklog_{id}" folder - its own
        // index.json and its photos and files together - then recomputes the workbook's status the
        // same way UpdateEntry does: removing the last still-Open worklog closes the workbook, and
        // removing the only one reopens it (a workbook with no worklogs is Open, per
        // RecomputeWorkbookStatus).
        //
        // ONE DELETE, NOT TWO. This used to remove the worklog's row from a shared entries.json and
        // then its attachment folder, in that order, precisely so a failure between the two could
        // not leave a surviving row pointing at photos that were gone. With the worklog's data
        // living in the same folder as its attachments that failure mode does not exist: the folder
        // either goes or it does not, and a failure leaves the worklog completely intact rather than
        // half-deleted. It is also exactly what a user does by hand in a file manager.
        //
        // The remaining worklogs deliberately KEEP their ids - nothing is renumbered. Ids are what
        // the board pills and the exported PDF show ("#4"), and they name the folders, so
        // renumbering would silently relabel every worklog the user has already written about
        // elsewhere. A gap in the numbering is the correct record of a deleted worklog.
        // ###########################################################################################
        public static bool DeleteEntry(int workbookId, int entryId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                Logger.Warning($"Failed to delete worklog: workbook [#{workbookId}] folder not found");
                return false;
            }

            var entries = ReadEntries(folder);
            int index = entries.FindIndex(e => e.Id == entryId);
            if (index < 0)
            {
                Logger.Warning($"Failed to delete worklog: worklog [#{entryId}] not found in workbook [#{workbookId}]");
                return false;
            }

            // The non-creating resolver: GetEntryAttachmentsFolder would re-create the very folder
            // being removed.
            string worklogFolder = Path.Combine(folder, BuildEntryAttachmentsFolderName(entryId));
            if (Directory.Exists(worklogFolder))
            {
                try
                {
                    Directory.Delete(worklogFolder, recursive: true);
                }
                catch (Exception ex)
                {
                    // The worklog is still entirely on disk, so it must still be reported as
                    // present - telling the caller it was deleted would leave the UI hiding a
                    // worklog that comes back on the next refresh.
                    Logger.Warning($"Failed to delete worklog [#{entryId}]: could not remove its folder [{worklogFolder}]: [{ex.Message}]");
                    return false;
                }
            }

            entries.RemoveAt(index);

            RecomputeWorkbookStatus(folder, workbookId, entries);

            Logger.Info($"Setting changed: [Worklog] deleted worklog [#{entryId}] from workbook [#{workbookId}]");

            return true;
        }

        // ###########################################################################################
        // Resolves (creating if missing) the subfolder that IS one worklog - its own index.json
        // plus its photo/file attachment bytes - named "worklog_<id>" inside its workbook's folder.
        // Photo/file metadata (comment, display order) lives in that index.json via
        // WorklogAttachmentRecord; the actual bytes sit beside it, named by
        // WorklogAttachmentRecord.FileName.
        // Returns null when the workbook folder cannot be found.
        // ###########################################################################################
        // ###########################################################################################
        // The NAME (not path) of one entry's attachment folder: "worklog_{id}".
        //
        // ONE definition, because this string was written out four times - the two resolvers below,
        // the draft-move helper, and its log lines - and any rename had to be made correctly in all
        // of them or attachments would be written to one folder and read from another.
        //
        // Named for the WORKLOG rather than the "entry" the code calls it internally: this folder is
        // visible to the user, both in the Workbooks folder on disk and inside an exported ZIP, and
        // the app says "worklog" everywhere a user can see one. It was "entry-{id}-files".
        // Deliberately NO migration of existing data - a folder from an older build keeps its name
        // and its attachments simply stop being found, which was accepted when this was requested.
        // ###########################################################################################
        public static string BuildEntryAttachmentsFolderName(int entryId) =>
            $"worklog_{entryId.ToString(CultureInfo.InvariantCulture)}";

        public static string? GetEntryAttachmentsFolder(int workbookId, int entryId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                return null;
            }

            string attachmentsFolder = Path.Combine(folder, BuildEntryAttachmentsFolderName(entryId));
            Directory.CreateDirectory(attachmentsFolder);
            return attachmentsFolder;
        }

        // ###########################################################################################
        // The same path as GetEntryAttachmentsFolder, but WITHOUT creating anything - for callers
        // that are about to DELETE the folder or merely want to know whether it exists.
        //
        // Resolving through the creating form for those is self-defeating: it re-creates the very
        // folder a caller is about to remove, so a delete leaves an empty folder behind where there
        // had been none. Returns null when the workbook folder cannot be found; the attachments
        // folder itself may or may not exist, which is the caller's business to check.
        // ###########################################################################################
        public static string? GetEntryAttachmentsFolderPath(int workbookId, int entryId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                return null;
            }

            return Path.Combine(folder, BuildEntryAttachmentsFolderName(entryId));
        }

        // ###########################################################################################
        // Whether the given entry id is currently used by a SAVED entry of the workbook.
        //
        // The editor's draft asks this before deleting the folder its reserved id names: a peeked
        // id is not a reservation, so by the time a draft is cancelled that number can legitimately
        // belong to a worklog saved meanwhile, whose record, photos and files all live in exactly
        // that folder. Deleting it then destroys that whole saved worklog.
        // ###########################################################################################
        public static bool EntryExists(int workbookId, int entryId)
        {
            string? folder = GetWorkbookFolder(workbookId);
            if (folder == null)
            {
                return false;
            }

            return ReadEntries(folder).Any(e => e.Id == entryId);
        }

        // ###########################################################################################
        // Recomputes and saves a workbook's entryCount and Open/Closed status from its current
        // entries: Closed once there is at least one entry and every one of them is Closed, Open
        // otherwise (including the no-entries case). Called after every entry is added, so a still
        // Open entry keeps (or reopens) the workbook, and closing the last outstanding entry closes
        // it.
        //
        // Also stamps EndDate on the TRANSITION into Closed - the date the job actually finished,
        // which the top bar shows in place of the start date. Only on the transition: recomputing a
        // workbook that is already Closed (any UpdateEntry on it, e.g. adding a photo to a closed
        // worklog) must not keep pushing the finish date forward to today. StartDate is never
        // written here at all.
        // ###########################################################################################
        private static void RecomputeWorkbookStatus(string folder, int workbookId, List<WorklogEntryRecord> entries)
        {
            // Re-read from disk rather than taking a record from the caller. That is load-bearing
            // now that index.json also carries LastWorklogId: AllocateNextEntryId has just written
            // the freshly allocated counter into this very file, and saving a record captured
            // before that would roll it straight back and re-hand the same id to the next entry.
            var record = ReadWorkbookRecord(folder);
            if (record == null)
            {
                return;
            }

            record.WorklogCount = entries.Count;

            bool wasOpen = IsWorkbookStatusOpen(record.Status);

            record.Status = entries.Count > 0 && entries.All(e => IsResolvedState(e.State))
                ? "Closed"
                : "Open";

            // The transition Open -> Closed, and only that. Re-closing an already-Closed workbook
            // is not a transition and leaves the original finish date standing.
            if (wasOpen && !IsWorkbookStatusOpen(record.Status))
            {
                record.EndDate = DateTime.Now.Date;
            }

            SaveWorkbook(folder, record);

            Logger.Info($"Setting changed: [Worklog] workbook [#{workbookId}] status set to [{record.Status}] ({entries.Count} worklogs)");
        }

        private static readonly JsonSerializerOptions SaveJsonOptions = new() { WriteIndented = true };

        // ###########################################################################################
        // Serializes ONE worklog and writes it to its own "worklog_{id}/index.json", creating that
        // folder when the worklog has no attachments yet, and returning false when nothing reached
        // disk. Callers MUST propagate that: AddEntry and UpdateEntry used to report success
        // regardless, so a locked or full disk closed the editor looking saved and the user then
        // watched their edits revert on the next refresh.
        //
        // One worklog per call, rather than rewriting every worklog in the workbook, is the point of
        // the per-folder layout: saving a worklog cannot touch - or lose - any other, and a failure
        // is confined to the one the user is editing.
        // ###########################################################################################
        private static bool SaveEntry(string workbookFolder, WorklogEntryRecord entry)
        {
            string worklogFolder = Path.Combine(workbookFolder, BuildEntryAttachmentsFolderName(entry.Id));

            try
            {
                Directory.CreateDirectory(worklogFolder);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to create worklog folder [{worklogFolder}]: {ex.Message}");
                return false;
            }

            string indexPath = Path.Combine(worklogFolder, AppConfig.WorklogIndexFileName);
            return AtomicJsonFile.Write(indexPath, entry, SaveJsonOptions, "worklog");
        }

        // ###########################################################################################
        // Serializes one workbook and writes it to its own index.json, returning false when nothing
        // reached disk.
        // ###########################################################################################
        private static bool SaveWorkbook(string folder, WorkbookRecord record)
        {
            string indexPath = Path.Combine(folder, AppConfig.WorklogIndexFileName);
            return AtomicJsonFile.Write(indexPath, record, SaveJsonOptions, "workbook");
        }
    }
}
