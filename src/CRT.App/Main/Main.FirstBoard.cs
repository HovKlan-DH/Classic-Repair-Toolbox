using Avalonia.Threading;
using System;
using System.Threading.Tasks;
using Tabs.TabSchematics;

namespace CRT
{
    // ###########################################################################################
    // *** THE FIRST BOARD IS DRAWN BEFORE THE WINDOW IS SHOWN (owner report, 2026-11-09: "It can take a couple of
    // seconds before it will populate the schematic image and thumbnails etc. and in the meantime I see an empty UI ...
    // Maybe the splash should show a little longer, until the UI has finished loading"). ***
    //
    // App keeps the splash up until FirstBoardShown says the first board load is over - its lists filled, its
    // thumbnails placed and its picture read - and only then shows this window, never waiting longer than
    // AppConfig.SplashWaitsForFirstBoardAtMost. The window loads unshown safely: the picture is fitted by layout
    // (Stretch="Uniform"), and the viewport clamps, labels and overlays itself again as its size first comes
    // (TabSchematics' Bounds handlers).
    //
    // Over whatever came of that load: no board listed, none selected, no data file, a file that could not be read,
    // no pictures, a picture that could not be read, a failure on the way. A load overtaken by another leaves it to
    // the newer one. Every board load reports here; only the first one's report changes anything.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        private readonly TaskCompletionSource thisFirstBoardShown = new(TaskCreationOptions.RunContinuationsAsynchronously);

        // The picture the first board shows, while its image is being read - null otherwise.
        private string? thisFirstPictureAwaited;

        // Done once the first board load is over, however it ended (above).
        internal Task FirstBoardShown => this.thisFirstBoardShown.Task;

        // ###########################################################################################
        // Waits for the first board to be drawn, at most `atMost`. True when it was; false when the wait gave up -
        // the window is shown all the same, and the board comes in as it is read.
        // ###########################################################################################
        internal async Task<bool> WaitForFirstBoardAsync(TimeSpan atMost)
        {
            Task first = this.FirstBoardShown;
            return await Task.WhenAny(first, Task.Delay(atMost)) == first;
        }

        // The first board load is over. On the UI thread.
        private void FirstBoardDone()
        {
            this.thisFirstPictureAwaited = null;
            this.thisFirstBoardShown.TrySetResult();
        }

        // The same, from the pool.
        private void FirstBoardDoneSoon() => Dispatcher.UIThread.Post(this.FirstBoardDone);

        // ###########################################################################################
        // The load's last step placed the thumbnails and is about to select `selected` - called BEFORE the selection,
        // as a picture with no file is settled inside it. With none there is nothing more to wait for.
        // ###########################################################################################
        private void FirstBoardPlaced(SchematicThumbnail? selected)
        {
            if (this.thisFirstBoardShown.Task.IsCompleted)
                return;

            if (selected == null)
                this.FirstBoardDone();
            else
                this.thisFirstPictureAwaited = selected.Name;
        }

        // The Schematics tab read a picture's image (or could not): the first board's is the last thing waited for.
        private void OnPictureSettled(string? picture)
        {
            if (this.thisFirstPictureAwaited != null &&
                string.Equals(picture?.Trim(), this.thisFirstPictureAwaited.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                this.FirstBoardDone();
            }
        }
    }
}
