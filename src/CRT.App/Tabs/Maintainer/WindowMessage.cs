using Avalonia.Controls;
using Avalonia.Media;

namespace CRT
{
    // ###########################################################################################
    // The one-line message under a window's controls - a refusal in red, anything else in grey,
    // hidden when there is nothing to say.
    //
    // ONE copy (code review, 2026-09-25). The Production, Maintainers and Unused files windows each
    // carried this body - four times, counting Production's second message line - so a change to
    // how a refusal is shown had to be made four times and would be missed in one. Each keeps a
    // one-line wrapper that names its own TextBlock and calls this - now the panels those windows
    // became (BetaView, UnusedFilesView, SystemView) and the main window's lists.
    //
    // UI code, not a Handlers/ class: it sets a control's properties and decides nothing worth a
    // test of its own - the same reason the windows' code-behind is verified by running the app.
    // ###########################################################################################
    internal static class WindowMessage
    {
        public static void Show(TextBlock? text, string? message, bool isError)
        {
            if (text is null)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = isError ? Brushes.IndianRed : Brushes.Gray;
        }
    }
}
