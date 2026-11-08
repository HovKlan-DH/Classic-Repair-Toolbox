using Avalonia.Input;
using Avalonia.Input.Platform;
using System;
using System.Linq;

namespace Handlers.Theming
{
    // ###########################################################################################
    // SHIFT+DELETE CUTS, IN EVERY TEXT BOX OF CRT (owner report, 2026-11-07: "when I mark some text in a textarea and I
    // press SHIFT+DELETE, then it normally cuts the text and copies it to clipboard - but it will not add it to
    // clipboard, when done inside app? CTRL+C works fine?"). Avalonia 12.1.4's keys list Ctrl+Insert as Copy and
    // Shift+Insert as Paste, but only Ctrl+X as Cut (PlatformHotkeyConfiguration, read decompiled), so Shift+Delete
    // reached the TextBox as a plain Delete: the selection went, and nothing was put on the clipboard.
    //
    // Added once at start-up to the platform's own list, which every TextBox asks at each key (TextBox.OnKeyDown), so
    // the table's cells, the drop-down lists' boxes and every dialog have it too. Not on macOS, where Delete is the
    // key a Windows keyboard calls Backspace and Shift+Delete means nothing of the kind.
    // ###########################################################################################
    internal static class TextEditingKeys
    {
        public static readonly KeyGesture ShiftDeleteCut = new(Key.Delete, KeyModifiers.Shift);

        // Shift+Delete added to the keys that cut - once, however often this is called.
        public static void UseShiftDeleteCut(PlatformHotkeyConfiguration? keys)
        {
            if (keys is null || OperatingSystem.IsMacOS())
            {
                return;
            }

            keys.Cut ??= [];

            if (!keys.Cut.Any(gesture => gesture.Key == ShiftDeleteCut.Key && gesture.KeyModifiers == ShiftDeleteCut.KeyModifiers))
            {
                keys.Cut.Add(ShiftDeleteCut);
            }
        }
    }
}
