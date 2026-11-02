using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // Hands a file the Maintainer tab saved for viewing to the operating system - a PDF opens in the
    // PDF viewer. ONE copy for the submission table's file card, the board table's and the file
    // tree's (code review, 2026-10-04: there were three identical ones), so a change to how such
    // files are opened is made once.
    //
    // *** THROUGH ExternalTargetLauncher, THE ONE WAY CRT OPENS A LOCAL FILE (code review, 2026-11-02). ***
    // It went through the window's own launcher, past that class's rules, with OpenedFiles' list as a
    // second one to keep in step. The root is the folder the file was saved under (OpenedFiles.Root),
    // so a path anywhere else is refused; OpenedFilesTests holds the two lists together. Its accepting
    // path starts a process, which no test may (test rule 6), so this line itself is not tested.
    //
    // UI-bound (a file is opened only for a control in a window), so it lives beside the tab.
    // ###########################################################################################
    internal static class MaintainerFileLauncher
    {
        // False when there is no window to launch from, or no program for the type - the caller
        // then says the file could not be opened.
        public static Task<bool> LaunchAsync(Visual? anchor, string fullPath) =>
            Task.FromResult(
                anchor is not null &&
                TopLevel.GetTopLevel(anchor) is not null &&
                ExternalTargetLauncher.TryOpen(fullPath, OpenedFiles.Root));
    }
}
