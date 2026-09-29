using System;
using System.Globalization;

namespace CRT
{
    // ###########################################################################################
    // What either application says when a wait ran into WaitLimit's two minutes (owner decision,
    // 2026-09-28: "it must be solid in validating if it did finish").
    //
    // In CRT.UI, shared, because it is VOCABULARY BOTH APPLICATIONS SHOW: the same limit must be
    // described the same way in CRT and in the Maintainer tab. Pure, so the sentences are tested.
    //
    // *** A TIMEOUT IS NEVER REPORTED AS A FAILURE. *** The server usually carries on after the
    // client stops listening, so "it failed, try again" would send a second publish after one that
    // worked. Each sentence says what is KNOWN: it finished, it has not (yet), or it could not be
    // checked - and in the last two, to look again before repeating it.
    // ###########################################################################################
    public static class WaitWording
    {
        // "2 minutes", read off the limit itself so the words cannot outlive a change to it.
        public static string Limit =>
            string.Create(CultureInfo.InvariantCulture, $"{(int)WaitLimit.Maximum.TotalMinutes} minutes");

        // The server stayed silent - for a READ, where nothing can have changed and trying again is
        // the whole answer.
        public static string NoAnswer => $"The server did not answer within {WaitWording.Limit}. Try again in a moment.";

        // ###########################################################################################
        // How long the overlay has been waiting, in words (owner request, 2026-09-28 - it was
        // "0:06"): "1 second elapsed", "10 seconds elapsed", "1 minute and 1 second elapsed",
        // "10 minutes and 59 seconds elapsed". A whole minute drops its seconds ("2 minutes
        // elapsed"), since "and 0 seconds" reads as padding. Nothing in the first second: the card
        // appears at 0.3 s, and "0 seconds elapsed" says nothing worth reading. Past an hour the
        // minutes keep counting ("75 minutes ...") - only a download reporting progress waits that
        // long, and it has its own percentage in the message.
        // ###########################################################################################
        public static string Elapsed(TimeSpan elapsed)
        {
            int total = (int)Math.Max(0, elapsed.TotalSeconds);

            if (total < 1)
                return string.Empty;

            int minutes = total / 60;
            int seconds = total % 60;

            string minutesPart = WaitWording.Count(minutes, "minute");
            string secondsPart = WaitWording.Count(seconds, "second");

            if (minutes == 0)
                return $"{secondsPart} elapsed";

            return seconds == 0
                ? $"{minutesPart} elapsed"
                : $"{minutesPart} and {secondsPart} elapsed";
        }

        private static string Count(int count, string unit) =>
            string.Create(CultureInfo.InvariantCulture, $"{count} {unit}{(count == 1 ? string.Empty : "s")}");

        // Shown on the overlay while a caller looks at the server again after a timeout.
        public static string Checking => $"The server did not answer within {WaitWording.Limit} - checking whether it finished...";

        // ###########################################################################################
        // After a timeout, and a look at the server: `finished` is what the look found - true, false,
        // or null when the look itself failed. `done` completes "it did finish: ..." and `notDone`
        // says what is still the case ("the submission is still waiting").
        // ###########################################################################################
        public static string AfterTimeout(bool? finished, string done, string notDone)
        {
            string silent = $"The server did not answer within {WaitWording.Limit}";

            return finished switch
            {
                true => $"{silent}, but it did finish: {WaitWording.Sentence(done)}",
                false => $"{silent}, and {WaitWording.Sentence(notDone)} It may still be working on it - wait a minute and look again before trying again.",
                _ => $"{silent}, and whether it finished could not be checked. Wait a minute and look again before trying again."
            };
        }

        // ###########################################################################################
        // Work on THIS machine that is still running at the limit (a PDF being written) - it cannot
        // be stopped halfway, so it carries on and is reported when it ends.
        // ###########################################################################################
        public static string StillRunning(string what) =>
            $"{what} is taking longer than {WaitWording.Limit}. It carries on in the background, and this says so when it is done.";

        // A clause made into a sentence: its trailing full stop, once.
        private static string Sentence(string clause)
        {
            string trimmed = clause.Trim();
            return trimmed.EndsWith('.') ? trimmed : trimmed + ".";
        }
    }
}
