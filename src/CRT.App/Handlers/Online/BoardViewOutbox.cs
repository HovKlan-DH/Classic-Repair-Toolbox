using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.OnlineHandling
{
    // The machine a report comes from, as the launch check-in describes it.
    public sealed record BoardViewMachine(string Version, string? OsHighlevel, string? OsVersion, string? Cpu);

    // ###########################################################################################
    // THE VIEWS WAITING TO BE SENT, AND THE ONE REPORT BEING SENT - kept in a small file so no view
    // is lost to a closed application or a missing network (owner: "must work robust, and take care
    // of troublesome or no network/internet").
    //
    //   Waiting - views counted and not yet in a report, oldest first.
    //   Sending - the report sent and not yet answered. It is sent AGAIN, UNCHANGED - the same
    //             BatchId and the same views - until an answer finishes it
    //             (BoardViewContract.DeliveryFor), so a report whose answer was lost is recognised
    //             by the server and never counted twice. Views counted meanwhile wait behind it.
    //
    // A view older than BoardViewRules.MaxAge is let go - the server would not count it - and past
    // MaxWaitingViews the oldest go, so a machine offline for months cannot grow the file for ever.
    //
    // Pure: BoardViewReporter reads and writes the file and talks to the server.
    // ###########################################################################################
    public sealed class BoardViewOutbox
    {
        private static readonly JsonSerializerOptions FileSettings = new(JsonSerializerDefaults.Web);

        public List<BoardView> Waiting { get; set; } = [];

        public BoardViewReport? Sending { get; set; }

        [JsonIgnore]
        public bool HasAnything => this.Sending is not null || this.Waiting.Count > 0;

        public void Add(BoardView view)
        {
            ArgumentNullException.ThrowIfNull(view);

            this.Waiting.Add(view);

            if (this.Waiting.Count > BoardViewRules.MaxWaitingViews)
                this.Waiting.RemoveRange(0, this.Waiting.Count - BoardViewRules.MaxWaitingViews);
        }

        // ###########################################################################################
        // The report to send now, or null when there is nothing worth sending: the one still being
        // sent if any of it is still countable, otherwise a new one of the oldest waiting views
        // (at most BoardViewRules.MaxViewsPerReport), which then becomes the one being sent.
        // ###########################################################################################
        public BoardViewReport? NextReport(DateTimeOffset now, BoardViewMachine machine)
        {
            ArgumentNullException.ThrowIfNull(machine);

            if (this.Sending is BoardViewReport sending)
            {
                if (sending.Views?.Any(view => view is not null && BoardViewRules.IsCountable(view.ViewedUtc, now)) == true)
                    return sending;

                // Too old to count, all of it: sending it again would only be refused.
                this.Sending = null;
            }

            this.Waiting.RemoveAll(view => view is null || !BoardViewRules.IsCountable(view.ViewedUtc, now));

            if (this.Waiting.Count == 0)
                return null;

            int take = Math.Min(BoardViewRules.MaxViewsPerReport, this.Waiting.Count);
            List<BoardView> views = this.Waiting.GetRange(0, take);
            this.Waiting.RemoveRange(0, take);

            this.Sending = new BoardViewReport(
                Guid.NewGuid().ToString("N"),
                machine.Version,
                machine.OsHighlevel,
                machine.OsVersion,
                machine.Cpu,
                views);

            return this.Sending;
        }

        // The report with this batch id got its answer - it is done with.
        public void Finished(string? batchId)
        {
            if (this.Sending is not null && string.Equals(this.Sending.BatchId, batchId, StringComparison.Ordinal))
                this.Sending = null;
        }

        public string ToJson() => JsonSerializer.Serialize(this, BoardViewOutbox.FileSettings);

        // A file that will not read is an empty outbox - it costs the views in it, never the
        // application.
        public static BoardViewOutbox FromJson(string? json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return new BoardViewOutbox();

            try
            {
                BoardViewOutbox? read = JsonSerializer.Deserialize<BoardViewOutbox>(json, BoardViewOutbox.FileSettings);

                if (read is null)
                    return new BoardViewOutbox();

                read.Waiting = read.Waiting?.Where(view => view is not null && !string.IsNullOrWhiteSpace(view.SystemId)).ToList() ?? [];
                return read;
            }
            catch (JsonException)
            {
                return new BoardViewOutbox();
            }
        }
    }
}
