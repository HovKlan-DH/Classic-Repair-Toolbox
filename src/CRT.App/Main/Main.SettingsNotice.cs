using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // "Your settings could not be read" (code review, 2026-10-13). UserSettings.LoadFrom keeps a
    // settings file it cannot use - moved aside as "<name>.corrupt-<date>", or left where it is -
    // and starts on defaults; this card says so, and where the file went, so a user whose
    // preferences all vanished at once is not left guessing.
    //
    // *** IT STAYS UNTIL CLOSED, although it has nothing to press (decided 2026-10-13, the owner
    // leaving it to what works best for the user). *** The one exception to "a card with nothing to
    // do goes by itself" (Main.Notices.cs): it appears at launch, while the first board loads, so
    // eight seconds can pass before anyone looks; it is the only place the kept file's name is
    // said; and when the file could not be moved aside, settings changed in this session are not
    // saved - something to know for as long as CRT runs. One of the cards in NoticeStack; see
    // Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Shows the card when the settings load at launch found the file unreadable. Called from
        // StartAsync - tests construct Main without it, and call this themselves.
        // ###########################################################################################
        internal void ShowSettingsUnreadableNoticeIfNeeded()
        {
            if (UserSettings.UnreadableSettings is not { } unreadable)
            {
                return;
            }

            this.ShowSettingsUnreadableNotice(unreadable.NoticeText);
        }

        private void ShowSettingsUnreadableNotice(string text)
        {
            NoticeEmphasis.Show(this.SettingsUnreadableBannerText, text);
            this.SettingsUnreadableBanner.IsVisible = true;
        }

        private void OnSettingsUnreadableBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.SettingsUnreadableBanner.IsVisible = false;
        }
    }
}
