using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;

namespace CRT
{
    // ###########################################################################################
    // WHAT AN OVERLAY DOES TO WHAT IT COVERS - shared by BusyOverlay (a whole window, while work
    // runs) and UpdateRequiredOverlay (a whole tab, for good):
    //
    //   - the other children of its host panel FADE (their opacity drops - a dark tint alone all but
    //     vanished in the dark theme), each remembered so it can be put back exactly;
    //   - keys are taken on the TUNNEL route, on the way down, before any control beneath sees them.
    //     A click cannot get past an overlay, but a key goes to whatever has focus - the button that
    //     started the work, still focused, would take Enter and start it again.
    //
    // One copy (code review, 2026-10-09: UpdateRequiredOverlay had re-implemented both), so a fix to
    // either reaches both overlays. A shared control's helper: it touches nothing of CRT's
    // orchestration.
    // ###########################################################################################
    internal static class OverlayCover
    {
        // What the covered controls fade to.
        internal const double Fade = 0.35;

        // ###########################################################################################
        // Fades every other child of `overlay`'s host panel not faded already, remembering its
        // opacity in `faded`. Nothing when the overlay is not in a panel.
        // ###########################################################################################
        internal static void FadeSiblings(Control overlay, IDictionary<Control, double> faded)
        {
            ArgumentNullException.ThrowIfNull(overlay);
            ArgumentNullException.ThrowIfNull(faded);

            if (overlay.Parent is not Panel host)
                return;

            foreach (Control sibling in host.Children)
            {
                if (ReferenceEquals(sibling, overlay) || faded.ContainsKey(sibling))
                    continue;

                faded[sibling] = sibling.Opacity;
                sibling.Opacity = OverlayCover.Fade;
            }
        }

        // Puts back exactly the opacity each faded control had, and forgets them.
        internal static void RestoreFaded(IDictionary<Control, double> faded)
        {
            ArgumentNullException.ThrowIfNull(faded);

            foreach ((Control control, double opacity) in faded)
                control.Opacity = opacity;

            faded.Clear();
        }

        // ###########################################################################################
        // Takes every key and every typed character meant for anything under `target`, on the
        // tunnel route - except one meant for `except` or a control inside it (the overlay's own
        // button). Disposing gives them back.
        // ###########################################################################################
        internal static IDisposable TakeKeys(Interactive target, Visual? except = null)
        {
            ArgumentNullException.ThrowIfNull(target);

            bool IsForExcepted(object? source) =>
                except is not null && source is Visual visual && (ReferenceEquals(visual, except) || except.IsVisualAncestorOf(visual));

            void OnKey(object? sender, KeyEventArgs e)
            {
                if (!IsForExcepted(e.Source))
                    e.Handled = true;
            }

            void OnText(object? sender, TextInputEventArgs e)
            {
                if (!IsForExcepted(e.Source))
                    e.Handled = true;
            }

            EventHandler<KeyEventArgs> key = OnKey;
            EventHandler<TextInputEventArgs> text = OnText;

            target.AddHandler(InputElement.KeyDownEvent, key, RoutingStrategies.Tunnel, handledEventsToo: true);
            target.AddHandler(InputElement.KeyUpEvent, key, RoutingStrategies.Tunnel, handledEventsToo: true);
            target.AddHandler(InputElement.TextInputEvent, text, RoutingStrategies.Tunnel, handledEventsToo: true);

            return new KeysTaken(target, key, text);
        }

        private sealed class KeysTaken(Interactive target, EventHandler<KeyEventArgs> key, EventHandler<TextInputEventArgs> text) : IDisposable
        {
            private bool thisGivenBack;

            public void Dispose()
            {
                if (this.thisGivenBack)
                    return;

                this.thisGivenBack = true;
                target.RemoveHandler(InputElement.KeyDownEvent, key);
                target.RemoveHandler(InputElement.KeyUpEvent, key);
                target.RemoveHandler(InputElement.TextInputEvent, text);
            }
        }
    }
}
