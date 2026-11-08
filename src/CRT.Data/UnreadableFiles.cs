using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** A FILE THAT CANNOT BE READ IS SAID ONCE, NOT AT EVERY READ OF IT (code review, 2026-11-07). *** A workbook's, a
    // worklog's or a device's index.json that cannot be read is skipped - the repair or the device is missing from the
    // lists - and that is an ERROR a person must know of ("Something went wrong", ProblemsLog). But the lists are read again
    // at every board switch, worklog action and device pick (RefreshWorklogBar), so one damaged file logged as an error at
    // every read brought the card back as soon as it was closed, counted up for ever, and pushed every other line out of
    // the 500 ProblemsLog keeps.
    //
    // So the line is logged the first time a file is found unreadable AS IT IS NOW - its length and write time - and a
    // debug line at every read after, while it stays so. A file changed since (mended by hand, or broken anew) is said
    // again. For the run only: the next launch says it once more.
    //
    // A FOLDER is said once for the run (its stamp never moves): a folder in the drafts folder that could not be taken in
    // as a draft, or a CRT 3 draft not converted yet - both found again at every load of the board list (DraftManager;
    // code review, 2026-11-08).
    //
    // *** IN CRT.Data, SO A BOARD FILE IS SAID SO TOO (code review, 2026-11-08). *** A board file that cannot be read is
    // read again at every board load and every look at its draft (BoardDataReader, BoardFiles), and its error came back as
    // the card was closed. Logged through CrtLog, which CRT hands to its own Logger (AppLoggerAdapter). Each LEVEL is
    // remembered apart: a file's warning and the error of the reader that met it are two lines, each said once.
    // ###########################################################################################
    internal static class UnreadableFiles
    {
        private static readonly object Lock = new();

        // Each file said, by level and full path, with its stamp when it was said.
        private static readonly Dictionary<string, (long Length, long WriteTicks)> Said = new(StringComparer.Ordinal);

        // ###########################################################################################
        // Logs `message` about the file at `path` that could not be read: an error the first time the file is so, a debug
        // line while it stays as it was.
        // ###########################################################################################
        public static void Report(string path, string message) =>
            UnreadableFiles.Say(path, message, "error", CrtLog.Critical);

        // The same, said as a warning: something nobody has to act on at once, listed under "Errors and warnings".
        public static void Warn(string path, string message) =>
            UnreadableFiles.Say(path, message, "warning", CrtLog.Warning);

        private static void Say(string path, string message, string level, Action<string> first)
        {
            if (UnreadableFiles.IsNew(level, path))
            {
                first(message);
            }
            else
            {
                CrtLog.Debug(message);
            }
        }

        // Whether the file at `path`, as it is now, has not been said at `level` yet - and from now on it has.
        private static bool IsNew(string level, string path)
        {
            string key = level + "|" + UnreadableFiles.KeyOf(path);
            (long, long) stamp = UnreadableFiles.StampOf(path);

            lock (UnreadableFiles.Lock)
            {
                if (UnreadableFiles.Said.TryGetValue(key, out (long, long) said) && said == stamp)
                {
                    return false;
                }

                UnreadableFiles.Said[key] = stamp;
                return true;
            }
        }

        private static string KeyOf(string path)
        {
            try
            {
                return Path.GetFullPath(path);
            }
            catch (Exception)
            {
                return path ?? string.Empty;
            }
        }

        private static (long Length, long WriteTicks) StampOf(string path)
        {
            try
            {
                var file = new FileInfo(path);
                return file.Exists ? (file.Length, file.LastWriteTimeUtc.Ticks) : (-1, -1);
            }
            catch (Exception)
            {
                return (-1, -1);
            }
        }
    }
}
