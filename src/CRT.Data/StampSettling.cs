using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHEN A FILE'S LENGTH AND WRITE TIME MAY BE TRUSTED TO SAY "UNCHANGED" (code review, 2026-10-13).
    //
    // Every cache over a file here is keyed by its length and last write time. A write time is only
    // as fine as the clock that stamps it - on Windows that clock ticks every few milliseconds - so
    // two writes of the same length inside one tick (a rectangle moved from X=10 to X=30, saved
    // twice in quick succession) leave the same stamp on different bytes, and a cache would answer
    // with the first file's value. git calls such an entry "racily clean" and re-reads it; this is
    // the same rule: a stamp is trusted only once the write is old enough that a second write could
    // no longer share it. A file written in the last moments is read afresh - the few a user just
    // saved, never the thousands nobody touched.
    // ###########################################################################################
    internal static class StampSettling
    {
        public static readonly TimeSpan SettleTime = TimeSpan.FromSeconds(2);

        public static bool IsSettled(DateTime lastWriteUtc) =>
            DateTime.UtcNow - lastWriteUtc >= StampSettling.SettleTime;
    }
}
