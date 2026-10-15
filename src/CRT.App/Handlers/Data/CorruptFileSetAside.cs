using System;
using System.Globalization;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A local JSON file CRT could not parse is MOVED ASIDE, never written over (code review,
    // 2026-10-13): the settings file and the drawn traces were both reset to empty on a parse
    // failure, and the next save wrote the empty state over the only copy. Moved to
    // "<path>.corrupt-<yyyyMMdd-HHmmss>" ("-2", "-3", ... when that name is taken), so a user - or
    // whoever helps them - can still get at what it held.
    // ###########################################################################################
    internal static class CorruptFileSetAside
    {
        // ###########################################################################################
        // Moves `path` aside with `move` (File.Move unless a test hands another). Returns where it
        // went, or null when it could not be moved - the caller must then not write over it.
        // ###########################################################################################
        internal static string? SetAside(string path, DateTime now, Action<string, string>? move = null)
        {
            string stem = $"{path}.corrupt-{now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)}";

            try
            {
                string target = stem;
                int attempt = 1;

                while (File.Exists(target) || Directory.Exists(target))
                {
                    attempt++;
                    target = $"{stem}-{attempt}";
                }

                (move ?? ((from, to) => File.Move(from, to)))(path, target);
                return target;
            }
            catch (Exception ex)
            {
                Logger.Warning($"An unreadable file could not be moved aside: [{path}] [{ex.Message}]");
                return null;
            }
        }
    }
}
