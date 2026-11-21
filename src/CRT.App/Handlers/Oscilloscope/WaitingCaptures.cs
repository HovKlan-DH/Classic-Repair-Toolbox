using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // CAPTURES LEFT WAITING GO (code review, 2026-11-21). An oscilloscope capture waits in AppConfig.ScopeCaptureWaitFolder, as
    // "<32 hex digits>.png", until it is accepted into a draft or let go of - and "Capture from oscilloscope pin-by-pin" takes many.
    // One CRT ending with captures not accepted (a crash, the process killed), or a delete Windows refused while a picture was being
    // read, stayed there for good. Swept once a run, after launch: only files of that very name, directly in the folder, through no
    // linked folder, and only those older than `olderThan` - another CRT running may have captures waiting there now.
    // ###########################################################################################
    public static class WaitingCaptures
    {
        // How old a capture left waiting is before it is swept - no window keeps one waiting that long.
        public static readonly TimeSpan SweptAfter = TimeSpan.FromDays(1);

        private static readonly Regex CaptureName = new("^[0-9a-f]{32}\\.png$", RegexOptions.CultureInvariant);

        // The captures in `folder` last written before `nowUtc` less `olderThan` deleted - how many went.
        public static int Sweep(string folder, DateTime nowUtc, TimeSpan olderThan)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder) ||
                new DirectoryInfo(folder).LinkTarget is not null)
            {
                return 0;
            }

            int removed = 0;
            DateTime before = nowUtc - olderThan;

            foreach (FileInfo file in new DirectoryInfo(folder).EnumerateFiles().Where(file => CaptureName.IsMatch(file.Name)))
            {
                try
                {
                    if (file.LinkTarget is null && file.LastWriteTimeUtc < before)
                    {
                        file.Delete();
                        removed++;
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Held by something for now: the next run tries again.
                }
            }

            return removed;
        }
    }
}
