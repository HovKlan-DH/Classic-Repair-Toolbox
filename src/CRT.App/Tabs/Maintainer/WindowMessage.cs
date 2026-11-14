using Avalonia.Controls;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // The one-line message under a window's controls - a refusal in the theme's red, anything else in its muted grey,
    // hidden when there is nothing to say (IndianRed and Gray in every theme until 2026-11-12).
    //
    // ONE copy (code review, 2026-09-25). The Production, Maintainers and Unused files windows each
    // carried this body - four times, counting Production's second message line - so a change to
    // how a refusal is shown had to be made four times and would be missed in one. Each keeps a
    // one-line wrapper that names its own TextBlock and calls this - now the panels those windows
    // became (BetaView, UnusedFilesView, BoardDetailView) and the main window's lists.
    //
    // UI code, not a Handlers/ class: it sets a control's properties and decides nothing worth a
    // test of its own - the same reason the windows' code-behind is verified by running the app.
    // ###########################################################################################
    internal static class WindowMessage
    {
        public static void Show(TextBlock? text, string? message, bool isError) =>
            WindowMessage.ShowIn(text, message, isError ? "Text_Fail_Fg" : "Text_Muted_Fg");

        // The same for a message that may be good news: a refusal in red, anything else in the theme's green.
        public static void ShowOutcome(TextBlock? text, string? message, bool isError) =>
            WindowMessage.ShowIn(text, message, isError ? "Text_Fail_Fg" : "Text_Success_Fg");

        // The one body both say through (code review, 2026-11-13: it was written twice again).
        private static void ShowIn(TextBlock? text, string? message, string colourKey)
        {
            if (text is null)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            ThemeResources.Paint(text, TextBlock.ForegroundProperty, colourKey);
        }
    }
}
