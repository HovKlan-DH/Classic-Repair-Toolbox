using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** A BOARD KEY IN ITS TWO PARTS - ONE RULE (code review, 2026-11-07). *** A board key is "{hardware}|{board}"
    // (Main.GetCurrentBoardKey builds it from the drop-downs). MyDevicesScope, the device form and Main each split it their
    // own way - one trimmed the parts, one did not, one took a key with a part missing as a whole - so the scope, the form
    // and the cards could disagree about which hardware a device is of, and a change to the key's shape had three places
    // to be made in. The hardware is what comes before the first "|", the board what comes after it, each trimmed; a key
    // with no "|" is a hardware alone.
    // ###########################################################################################
    internal static class BoardKeyParts
    {
        public static (string Hardware, string Board) Split(string? boardKey)
        {
            string key = boardKey?.Trim() ?? string.Empty;
            int bar = key.IndexOf('|', StringComparison.Ordinal);

            return bar < 0 ? (key, string.Empty) : (key[..bar].Trim(), key[(bar + 1)..].Trim());
        }
    }
}
