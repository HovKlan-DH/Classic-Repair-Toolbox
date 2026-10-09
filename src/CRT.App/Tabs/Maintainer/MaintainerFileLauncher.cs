using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform.Storage;

namespace CRT
{
    // ###########################################################################################
    // Hands a file the Maintainer tab saved for viewing to the operating system - a PDF opens in the
    // PDF viewer. ONE copy for the submission table's file card, the board table's and the file
    // tree's (code review, 2026-10-04: there were three identical ones), so a change to how such
    // files are opened is made once.
    //
    // UI-bound (it needs the window's launcher), so it lives beside the tab, not in Handlers/.
    // ###########################################################################################
    internal static class MaintainerFileLauncher
    {
        // False when there is no window to launch from, or no program for the type - the caller
        // then says the file could not be opened.
        public static async Task<bool> LaunchAsync(Visual? anchor, string fullPath)
        {
            try
            {
                return anchor is not null &&
                       TopLevel.GetTopLevel(anchor)?.Launcher is { } launcher &&
                       await launcher.LaunchFileInfoAsync(new FileInfo(fullPath));
            }
            catch (Exception)
            {
                // A platform with no handler for the type.
                return false;
            }
        }
    }
}
