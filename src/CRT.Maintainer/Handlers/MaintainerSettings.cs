using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // WHAT THE MAINTAINER APPLICATION REMEMBERS BETWEEN RUNS (owner requests, 2026-09-26): where its
    // window was - "if it is maximized, then start again in maximized, and if it is windowed, then
    // start same position and size next time" - and whether "Show changes only" was ticked.
    //
    // A plain JSON file beside the remembered session (ReviewSessionStore.AppFolderName). Nothing in
    // it is a secret, so unlike the session it is not protected.
    //
    // *** A FILE THAT CANNOT BE READ IS DEFAULTS, NEVER AN ERROR. *** It is a convenience: a
    // half-written or hand-edited file must cost a window position, not a start-up.
    //
    // The window reads and writes it through MaintainerMain.UseSettings, which only the running
    // application calls - so no test ever reads or writes the user's real file.
    // ###########################################################################################
    public sealed class MaintainerSettings
    {
        public bool HasWindowPlacement { get; set; }

        public bool WindowMaximized { get; set; }

        // The NORMAL (un-maximized) position and size - what a maximized window returns to.
        public int WindowX { get; set; }

        public int WindowY { get; set; }

        public double WindowWidth { get; set; }

        public double WindowHeight { get; set; }

        // The top-left of the screen the window was on, so a maximized window maximizes on THAT
        // screen - setting Maximized alone maximizes on whichever monitor the system picks.
        public int ScreenX { get; set; }

        public int ScreenY { get; set; }

        public bool ShowChangesOnly { get; set; }
    }

    public static class MaintainerSettingsStore
    {
        public const string FileName = "CRT-Maintainer-Settings.json";

        private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

        public static string DefaultPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ReviewSessionStore.AppFolderName,
                MaintainerSettingsStore.FileName);

        public static MaintainerSettings Load(string path)
        {
            try
            {
                if (!File.Exists(path))
                    return new MaintainerSettings();

                return JsonSerializer.Deserialize<MaintainerSettings>(File.ReadAllText(path), MaintainerSettingsStore.JsonOptions)
                    ?? new MaintainerSettings();
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
            {
                return new MaintainerSettings();
            }
        }

        // Written to a temporary file and moved into place, so a crash mid-write leaves the old
        // file rather than half of a new one. A failure is swallowed - see the header.
        public static void Save(string path, MaintainerSettings settings)
        {
            ArgumentNullException.ThrowIfNull(settings);

            try
            {
                string? directory = Path.GetDirectoryName(path);

                if (!string.IsNullOrEmpty(directory))
                    Directory.CreateDirectory(directory);

                string temporary = path + ".tmp";
                File.WriteAllText(temporary, JsonSerializer.Serialize(settings, MaintainerSettingsStore.JsonOptions));
                File.Move(temporary, path, overwrite: true);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                // Nothing to do: next time starts where the window's defaults put it.
            }
        }
    }

    // ###########################################################################################
    // Whether a remembered window would open where it can be seen. A monitor unplugged or a
    // resolution lowered since would put it off every screen, where it cannot be moved back; its
    // centre must be on one, or the window opens centred on the main screen instead. The same rule
    // CRT's own main window restores by.
    // ###########################################################################################
    public static class WindowPlacementRules
    {
        public static bool IsCentreOnAScreen(
            int x,
            int y,
            double width,
            double height,
            double scaling,
            IEnumerable<(int X, int Y, int Width, int Height)> screens)
        {
            ArgumentNullException.ThrowIfNull(screens);

            double factor = scaling > 0 ? scaling : 1;
            int centreX = x + (int)(width * factor / 2);
            int centreY = y + (int)(height * factor / 2);

            return screens.Any(screen =>
                centreX >= screen.X &&
                centreY >= screen.Y &&
                centreX < screen.X + screen.Width &&
                centreY < screen.Y + screen.Height);
        }
    }
}
