using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Builds and validates the identity of a system created through "Add a new system"
    // (NewContributeStrategy.md Phase 2, session 2c, task 9) - the three folder segments the user
    // types (manufacturer, hardware, board) and the ExcelDataFile key derived from them.
    //
    // WHY ExcelDataFile NAMES A FILE THAT WILL NEVER EXIST. Everywhere else in the app, a system is
    // identified by its ExcelDataFile - it is the BoardDataReader cache key, the input to
    // DraftManager.GetSystemFolder, and what the worklog board key and schematic image paths are
    // resolved against. A draft-only system has no .xlsx at all (nothing in this phase writes one -
    // every edit goes into draft.json, see LabelEditorDraftWriter's header for that decision), but
    // it still needs that identity string, and the string has to be FOLDER-SHAPED because three
    // consumers read meaning out of its segments:
    //   - DraftManager.GetSystemFolder splits on "/" and drops the last segment, so the draft folder
    //     mirrors Data/'s own Manufacturer/Hardware/Board layout (needs at least 2 segments);
    //   - HardwareBoardEntry.ShortHardwareBoardLabel takes segments[^3]/segments[^2] (needs 3);
    //   - DraftFileResolver resolves a drafted file relative to it.
    // Three folder segments plus a file name satisfies all three.
    //
    // No version suffix, deliberately - real boards ship "Data C64 250407 v2.0.0.xlsx", where the
    // version tracks the resolved MAIN workbook's version. Keeping that in step by hand is exactly
    // one of the manual burdens task 9 exists to remove, and nothing in the app parses it.
    //
    // Pure string work - no file I/O, no Avalonia - so it is unit tested directly.
    // ###########################################################################################
    public static class NewSystemIdentity
    {
        // Long enough for any real manufacturer, hardware or board name, short enough that three of
        // them plus a file name stay clear of path-length limits on every platform.
        public const int MaxSegmentLength = 64;

        // Windows refuses these as file OR folder names whatever the extension, and a folder that
        // cannot be created would leave a registered system pointing at nothing. Checked on every
        // platform, not just Windows: a draft folder created on Linux is meant to be readable after
        // the same AppData folder is carried to a Windows machine.
        private static readonly HashSet<string> ReservedDeviceNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
        };

        // ###########################################################################################
        // The ExcelDataFile identity for a new system:
        // "<Manufacturer>/<Hardware>/<Board>/Data <Hardware> <Board>.xlsx".
        //
        // Always "/" as the separator, never Path.DirectorySeparatorChar - ExcelDataFile follows the
        // sync manifest's own convention (see DataManager), and a "\" here would not match the split
        // in DraftManager.GetSystemFolder or HardwareBoardEntry.ShortHardwareBoardLabel on Windows.
        //
        // Returns an empty string when any segment is blank, so a caller that skipped validation
        // cannot silently produce a half-formed key like "Commodore//250407/Data  250407.xlsx".
        // ###########################################################################################
        public static string BuildExcelDataFile(string manufacturer, string hardware, string board)
        {
            string cleanManufacturer = SanitizePathSegment(manufacturer);
            string cleanHardware = SanitizePathSegment(hardware);
            string cleanBoard = SanitizePathSegment(board);

            if (cleanManufacturer.Length == 0 || cleanHardware.Length == 0 || cleanBoard.Length == 0)
            {
                return string.Empty;
            }

            return $"{cleanManufacturer}/{cleanHardware}/{cleanBoard}/Data {cleanHardware} {cleanBoard}.xlsx";
        }

        // ###########################################################################################
        // Trims and collapses a typed segment into what will become a real folder name. Deliberately
        // NOT a character-stripping "make anything work" pass: an illegal character is reported by
        // IsValidPathSegment and corrected by the user, because silently turning "C64:NTSC" into
        // "C64NTSC" would create a folder the user did not ask for and cannot find later.
        //
        // What it does do is normalize whitespace (outer trim, inner runs collapsed to one space),
        // which is invisible to the user and would otherwise make "C 64" and "C  64" two different
        // systems.
        // ###########################################################################################
        public static string SanitizePathSegment(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        }

        // ###########################################################################################
        // Whether one typed segment can become a folder name, with a reason fit to show the user
        // when it cannot. The rules exist because each failure mode produces a system that looks
        // created but is broken in a way the user cannot diagnose:
        //   - blank: no folder to create;
        //   - invalid characters: Directory.CreateDirectory throws, or (on "/" and "\") silently
        //     creates a deeper tree that GetSystemFolder's segment count no longer matches. The
        //     set checked is the union across EVERY platform, not the host's - see the comment on
        //     invalidChars below;
        //   - "." / "..": resolve to a different folder entirely;
        //   - reserved device name: Windows refuses the folder outright;
        //   - trailing dot or space: Windows SILENTLY STRIPS them, so the folder on disk no longer
        //     matches the ExcelDataFile key stored in draft.json, and the draft becomes unfindable;
        //   - over-length: risks the platform path limit once combined with the drafts root.
        // ###########################################################################################
        public static bool IsValidPathSegment(string? value, out string reason)
        {
            string trimmed = SanitizePathSegment(value);

            if (trimmed.Length == 0)
            {
                reason = "cannot be blank";
                return false;
            }

            if (trimmed.Length > MaxSegmentLength)
            {
                reason = $"is too long (maximum {MaxSegmentLength} characters)";
                return false;
            }

            // *** THE INVALID SET IS THE UNION ACROSS EVERY PLATFORM, NOT THIS ONE'S. ***
            //
            // Path.GetInvalidFileNameChars is HOST-SPECIFIC: on Windows it returns the full set
            // below, while on Linux it returns only "/" and NUL. Trusting it would mean a system
            // created on Linux could be named "Commo*dore" - accepted there, and then impossible
            // to sync for every Windows user who downloads it, with the failure appearing on a
            // machine that never saw the name being typed.
            //
            // These names become FOLDER SEGMENTS in a tree that every platform syncs, so the rule
            // has to be the strictest platform's. WorkbookExportModel.SanitizeForFileName already
            // documents this exact reasoning for exported file names; this is the same lesson
            // applied to the segment that becomes a directory.
            //
            // The set is written out rather than derived, so it cannot change under the code when
            // the host does. Path.GetInvalidFileNameChars is still unioned in, in case a platform
            // refuses something not listed here.
            var invalidChars = new HashSet<char>(Path.GetInvalidFileNameChars())
            {
                // Windows' reserved set. "/" and NUL are the only two Linux refuses, and both are
                // already here.
                '<', '>', ':', '"', '/', '\\', '|', '?', '*', '\0'
            };

            // Control characters are refused everywhere and would break a log line or a manifest
            // entry even where the filesystem tolerated them.
            for (char control = (char)1; control < (char)32; control++)
                invalidChars.Add(control);

            foreach (char candidate in trimmed)
            {
                if (invalidChars.Contains(candidate))
                {
                    reason = $"cannot contain [{candidate}]";
                    return false;
                }
            }

            if (trimmed == "." || trimmed == "..")
            {
                reason = "cannot be [.] or [..]";
                return false;
            }

            // A reserved name is refused with or without an extension, so test the part before the
            // first dot as well as the whole string.
            string beforeExtension = trimmed.Split('.')[0];
            if (ReservedDeviceNames.Contains(trimmed) || ReservedDeviceNames.Contains(beforeExtension))
            {
                reason = $"cannot be [{trimmed}] - that name is reserved by the operating system";
                return false;
            }

            if (trimmed.EndsWith('.'))
            {
                reason = "cannot end with a dot";
                return false;
            }

            reason = string.Empty;
            return true;
        }

        // ###########################################################################################
        // The manufacturer segment of an existing system's ExcelDataFile
        // ("Commodore/C64/250407/Data ....xlsx" -> "Commodore"), so the create dialog can offer the
        // manufacturers already in use rather than inviting a fresh typo ("Comodore/") that would
        // silently create a second top-level folder beside the real one.
        //
        // Returns empty for a key with no folder segments at all, which the caller filters out.
        // ###########################################################################################
        public static string ExtractManufacturer(string? excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(excelDataFile))
            {
                return string.Empty;
            }

            string[] segments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);
            return segments.Length < 2 ? string.Empty : segments[0];
        }

        // ###########################################################################################
        // The HARDWARE and BOARD segments of an existing system's ExcelDataFile
        // ("Commodore/C64/250407/Data ....xlsx" -> "C64" and "250407").
        //
        // *** THESE ARE PATH SEGMENTS, NOT DISPLAY NAMES, and the difference is what made a
        // submission fail (2026-09-23). *** TabDrafts built the submission's SystemId from the
        // path ("Commodore/C64/250407") while sending Hardware and Board as the master workbook's
        // display names ("Commodore 64" and "250407 (long board)"). SubmissionValidator rebuilds
        // the id from those three parts and compares, so it correctly reported that the identifier
        // did not match the parts - and the message ("a fault in the submitting application") was
        // exactly right.
        //
        // The server's `systems` table stores the three parts alongside system_id precisely so the
        // Maintainer tab can list by them without parsing the id, which only works while they ARE the
        // id's parts. Display names belong on screen, not in the identity.
        // ###########################################################################################
        public static string ExtractHardware(string? excelDataFile)
        {
            return NewSystemIdentity.SegmentAt(excelDataFile, 1);
        }

        public static string ExtractBoard(string? excelDataFile)
        {
            return NewSystemIdentity.SegmentAt(excelDataFile, 2);
        }

        // Requires a FOLDER segment at the index: the last segment of an ExcelDataFile is the file
        // name, so a key must have more segments than the index being read for that index to name
        // a folder rather than the file itself.
        private static string SegmentAt(string? excelDataFile, int index)
        {
            if (string.IsNullOrWhiteSpace(excelDataFile))
            {
                return string.Empty;
            }

            string[] segments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);

            return segments.Length <= index + 1 ? string.Empty : segments[index];
        }
    }
}
