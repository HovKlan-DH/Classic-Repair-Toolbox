using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.Geometry;

namespace Handlers.Theming
{
    // ###########################################################################################
    // A DIALOG THAT SIZES ITSELF TO ITS CONTENT, KEPT ON THE SCREEN IT OPENS ON (DialogPlacement):
    //
    //   - NO TALLER THAN THE SCREEN'S WORKING AREA (MaxHeight), so its content scrolls and its buttons stay on screen -
    //     the dialog's markup puts what can grow in a ScrollViewer and the buttons in a row of their own under it.
    //   - CENTRED OVER ITS OWNER WHENEVER ITS SIZE CHANGES: a window sized to its content keeps its top edge where it
    //     was as it shrinks or grows.
    //
    // The Submit dialog's since the owner's report of 2026-11-01 ("I cannot see the submit button"), and "Add a new
    // board"'s since the owner's report of 2026-11-25 about the "Maintainer" tab's sign-in out of reach on Linux - the
    // search that followed measured that dialog 732 high as made, past the bottom of a 768 screen with nothing to scroll.
    // ###########################################################################################
    internal static class ScreenSizedDialog
    {
        public static void Use(Window dialog)
        {
            dialog.Opened += (_, _) =>
            {
                if (dialog.Screens.ScreenFromWindow(dialog) is { } screen)
                {
                    dialog.MaxHeight = DialogPlacement.MaxHeight(screen.WorkingArea.Height, screen.Scaling);
                }

                ScreenSizedDialog.CentreOnOwner(dialog);
            };

            dialog.PropertyChanged += (_, change) =>
            {
                if (change.Property == TopLevel.ClientSizeProperty && dialog.IsVisible)
                {
                    Dispatcher.UIThread.Post(() => ScreenSizedDialog.CentreOnOwner(dialog));
                }
            };
        }

        private static void CentreOnOwner(Window dialog)
        {
            if (dialog.Owner is not Window owner || dialog.FrameSize is not { } frame || owner.FrameSize is not { } ownerFrame)
            {
                return;
            }

            double ownerScaling = owner.DesktopScaling;
            var ownerBounds = new PixelRect(
                owner.Position,
                new PixelSize((int)Math.Round(ownerFrame.Width * ownerScaling), (int)Math.Round(ownerFrame.Height * ownerScaling)));

            Avalonia.Platform.Screen? screen = dialog.Screens.ScreenFromWindow(owner);

            PixelPoint wanted = DialogPlacement.CentredOver(ownerBounds, frame, dialog.DesktopScaling, screen?.WorkingArea ?? default);

            if (wanted != dialog.Position)
            {
                dialog.Position = wanted;
            }
        }
    }
}
