using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // BOARD VIEWS: WHICH BOARDS ARE LOOKED AT, AND WHERE (owner request, 2026-09-27: "whenever a
    // user selects a new board from the drop-down list, then hardware/board should be sent to my
    // backend server ... The user must not be identified - the goal here is only knowing usage on
    // boards", then "I want it to catch every time a user selects a board ... count only when
    // viewed for +10 seconds", and "a new table that can hold all this data in one table, as I then
    // can see usage of boards in countries").
    //
    // CRT counts a VIEW when a published board has been on screen for MinimumTimeOnScreen, EVERY
    // time - not once a day. Views wait in a small file on the user's machine and are sent to
    // CRT.Server in batches; the server stores ONE ROW PER VIEW in `crt_board_views` with the
    // board's names, the CRT version, the operating system, the CPU and the COUNTRY - looked up
    // from the sender's address as it arrives, and the address then forgotten. Nothing identifies a
    // user: no address, no installation id, nothing that links one view to another.
    //
    // *** MANDATORY, like the launch check-in (owner decision, 2026-09-27: "There will be no
    // opt-out option ... It is just alike 'ping-home'"). *** The Wiki's "Information collected"
    // page says what is sent.
    //
    // *** THIS FILE IS THE CONTRACT BOTH ENDS COMPILE AGAINST. *** CRT builds a BoardViewReport and
    // sends ToJson(report); the server binds the same record with the same settings
    // (ReviewApiContract.WireSettings) at the same route (PathUnderApi). BoardViewRules is what both
    // sides apply - CRT drops what the server would refuse rather than retrying it for ever.
    // ###########################################################################################
    public static class BoardViewContract
    {
        // The route under the API base: CRT posts to "<CrtServerBaseUrl>/usage/board-views", the
        // server maps "/api/usage/board-views".
        public const string PathUnderApi = "usage/board-views";

        // The body CRT sends. The server reads it with ReviewApiContract.WireSettings too.
        public static string ToJson(BoardViewReport report)
        {
            ArgumentNullException.ThrowIfNull(report);

            return JsonSerializer.Serialize(report, ReviewApiContract.WireSettings);
        }

        // ###########################################################################################
        // What CRT does with a report after the server answered `status` (0 = no answer at all).
        //
        // Accepted (2xx) is DONE, and so is a REFUSAL OF THIS REPORT - 400 (the flow refused it),
        // 413 (too large), 415 (not JSON), 422 (unprocessable): sending it again would be refused
        // again, and keeping it would block every view behind it. EVERYTHING ELSE keeps it - no
        // answer, a timeout, too many (429), a server fault (5xx) - and it is then sent again
        // unchanged, with the same BatchId, so the server can tell a retry from a new report and
        // never counts one view twice.
        //
        // *** ANY OTHER 4xx IS "NOT YET", NOT "NO" (2026-09-27). *** 404 is a server without the
        // route (a CRT released before its server, or Apache not proxying the path); 401, 403, 405
        // and the like are a proxy or firewall set up wrong on the server's side. None of them says
        // anything about the report, and dropping on them would make every CRT throw its views away
        // until the server was put right. Kept, the views wait - at most BoardViewRules.MaxAge, and
        // at most MaxWaitingViews of them - so a problem that is never fixed still ends.
        // ###########################################################################################
        public static BoardViewDelivery DeliveryFor(int status) =>
            status switch
            {
                >= 200 and < 300 => BoardViewDelivery.Done,
                400 or 413 or 415 or 422 => BoardViewDelivery.Done,
                _ => BoardViewDelivery.TryLater
            };
    }

    public enum BoardViewDelivery
    {
        Done,
        TryLater
    }

    // ###########################################################################################
    // One report: a batch of views from one machine.
    //
    //   BatchId  - a new random id per batch (never per user): a retry carries the same one, so the
    //              server stores each batch once. It links the views of ONE batch only, and the
    //              server keeps it apart from the views, for 60 days, purely to spot retries.
    //   Version  - "CRT 2026.10.0", the form crt_update.version has.
    //   OsHighlevel / OsVersion / Cpu - what the launch check-in sends.
    //
    // Nullable throughout because it is bound from a request anybody can send; BoardViewRules and
    // the server's flow decide what a usable one is.
    // ###########################################################################################
    public sealed record BoardViewReport(
        string? BatchId,
        string? Version,
        string? OsHighlevel,
        string? OsVersion,
        string? Cpu,
        IReadOnlyList<BoardView>? Views);

    // One view: the board ("Commodore/C64/250407", BoardDescriptorRules' id), when it had been on
    // screen long enough to count (UTC), and whether CRT was downloading from the BETA source.
    //
    // SystemId is what BoardId was called before "system" became "board" everywhere (owner
    // decision, 2026-10-09). Nothing writes it any more, but a CRT built before that still sends
    // it, and an outbox file one of them wrote still holds it - so both ends read IdOf, never
    // BoardId alone.
    public sealed record BoardView(
        string? BoardId,
        DateTimeOffset ViewedUtc,
        bool FromBeta,
        [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? SystemId = null)
    {
        public static string? IdOf(BoardView? view) =>
            string.IsNullOrWhiteSpace(view?.BoardId) ? view?.SystemId : view.BoardId;
    }

    // ###########################################################################################
    // The rules both ends apply.
    // ###########################################################################################
    public static class BoardViewRules
    {
        // How long a board must stay on screen before it counts as viewed (owner decision,
        // 2026-09-27) - long enough that flicking through the drop-down counts nothing.
        public static readonly TimeSpan MinimumTimeOnScreen = TimeSpan.FromSeconds(10);

        // The most views one report may carry. CRT sends a longer backlog as several reports.
        public const int MaxViewsPerReport = 200;

        // How old a view may be and still be counted - a machine offline for longer loses the
        // oldest. Views dated in the future by a fast clock are accepted within MaxAhead (and stored
        // at the server's time), beyond it refused.
        public static readonly TimeSpan MaxAge = TimeSpan.FromDays(30);
        public static readonly TimeSpan MaxAhead = TimeSpan.FromDays(1);

        // The most views CRT keeps waiting to be sent; beyond it the oldest go. At one view per
        // ten seconds that is a fortnight of a board changing every minute, all day.
        public const int MaxWaitingViews = 5000;

        // How many reports one address may send in an hour. CRT sends at most about one a minute
        // while someone is changing boards, so a household of several users fits; a script
        // inflating the numbers is held back. A report refused for it is sent again later.
        public const int MaxReportsPerAddressPerHour = 60;

        // Column widths in crt_board_views; longer values are cut to fit.
        public const int BoardIdLength = 200;
        public const int NameLength = 100;
        public const int VersionLength = 100;
        public const int OsHighlevelLength = 20;
        public const int OsVersionLength = 200;
        public const int CpuLength = 50;

        // Is a view from `viewedUtc` still counted at `now`?
        public static bool IsCountable(DateTimeOffset viewedUtc, DateTimeOffset now) =>
            viewedUtc >= now - BoardViewRules.MaxAge && viewedUtc <= now + BoardViewRules.MaxAhead;

        // The time a view is stored at: its own, but never later than the server's clock, and to
        // the second.
        public static DateTimeOffset StoredTime(DateTimeOffset viewedUtc, DateTimeOffset now)
        {
            DateTimeOffset at = (viewedUtc > now ? now : viewedUtc).ToUniversalTime();

            return new DateTimeOffset(at.Ticks - at.Ticks % TimeSpan.TicksPerSecond, TimeSpan.Zero);
        }

        // ###########################################################################################
        // A sender's text made fit to store: control characters removed, trimmed, cut to `maxLength`.
        // Null for nothing left. These values reach a public web page (Fun facts), which escapes them
        // too - this is the first of two guards, not the only one.
        //
        // *** NEVER HALF A CHARACTER (code review, 2026-09-29). *** `maxLength` counts UTF-16 code
        // units, so cutting at it could split a character outside the Basic Multilingual Plane (an
        // emoji, some CJK) and leave a lone high surrogate - which MariaDB's utf8mb4 refuses, failing
        // the WHOLE batch's insert. CRT then re-sent that batch every 15 minutes for ever, with every
        // later view queued behind it. So the cut steps back off a high surrogate, and a lone
        // surrogate already in the text (half a character from a broken sender) is dropped.
        // ###########################################################################################
        public static string? Clip(string? value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            var text = new StringBuilder(value.Length);

            for (int index = 0; index < value.Length; index++)
            {
                char character = value[index];

                if (char.IsHighSurrogate(character))
                {
                    // Kept only with its low half, as the pair it is.
                    if (index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]))
                    {
                        text.Append(character).Append(value[index + 1]);
                        index++;
                    }

                    continue;
                }

                if (char.IsLowSurrogate(character) || char.IsControl(character))
                    continue;

                text.Append(character);
            }

            string clean = text.ToString().Trim();

            if (clean.Length > maxLength)
            {
                int cut = maxLength;

                // Not between the two halves of one character.
                if (cut > 0 && char.IsHighSurrogate(clean[cut - 1]))
                    cut--;

                clean = clean[..cut].TrimEnd();
            }

            return clean.Length == 0 ? null : clean;
        }
    }

    // ###########################################################################################
    // One board's views, for the Maintainer tab's Boards screen (BoardDetailAnswer.Views). Views
    // from the BETA download source - mostly maintainers checking their own work - are left out of
    // every count but FromBetaLast30Days, so they cannot make a board look used.
    //
    // The windows are whole UTC days including today: "the last 7 days" is today and the six before.
    // TopCountries: the last 365 days, most views first, at most five.
    //
    // Daily (owner request, 2026-10-09: "a graph for showing usage of board per day"): the views of
    // each UTC day of the last 365 that had any, oldest first - a day with none is not listed, and
    // the Statistics view draws it at 0. Null from a server older than that.
    // ###########################################################################################
    public sealed record BoardViewStatistics(
        int Last7Days,
        int Last30Days,
        int Last365Days,
        int FromBetaLast30Days,
        IReadOnlyList<BoardViewCountry> TopCountries,
        IReadOnlyList<BoardViewDay>? Daily = null);

    // One day's views - a UTC day, as every window here is.
    public sealed record BoardViewDay(DateOnly Day, int Views);

    public sealed record BoardViewCountry(string CountryCode, string CountryName, int Views);
}
