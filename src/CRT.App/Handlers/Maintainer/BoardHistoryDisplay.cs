using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // A BOARD'S HISTORY VIEW (owner request, 2026-10-04: "another 'History' tab/button, after the
    // 'Maintainer' button ... it should show the full history of what has happened with this board,
    // in an 'easy to overview' way... which is not what the current history is").
    //
    // What made the old list hard to read was that it was one line per EVENT: a submission's sending
    // and its decision were two lines apart, with other submissions and pool changes between them,
    // and nothing said what any of them had changed. So:
    //
    //   - ONE CARD PER SUBMISSION, holding everything that happened to it in the order it happened
    //     (sent, changed by a maintainer, decided, its draft discarded), its state now in CRT's own
    //     words, what the contributor was told - and WHAT IT CHANGED, from the summary the server
    //     recorded as it went into BETA (CRT.Data's SubmissionChanges).
    //   - Everything else - a publish to stable, a push-back, a pool change, a placement - one line
    //     each, as before (BoardsDisplay.HistoryWhat / HistoryFooter).
    //   - Newest first, under a heading per MONTH, each item with its day in a column of its own, so
    //     the left edge reads as a timeline. A card sits at its LATEST event: a submission sent in
    //     September and published in October is October's.
    //
    // Pure, so the grouping and every word are tested; BoardDetailView.History.cs only draws it.
    // ###########################################################################################
    public static class BoardHistoryDisplay
    {
        // The submission events a card gathers - the rest stay lines of their own.
        private static readonly HashSet<string> CardEvents = new(StringComparer.Ordinal)
        {
            BoardHistoryEvents.Sent,
            BoardHistoryEvents.Decided,
            BoardHistoryEvents.Amended,
            BoardHistoryEvents.DraftDiscarded
        };

        // ###########################################################################################
        // The whole view, newest month first. A submission the server listed is a card even when no
        // event names it; an event naming a submission it did not list is a line of its own, as
        // before - nothing the server sent is dropped.
        // ###########################################################################################
        public static IReadOnlyList<HistoryMonth> Build(BoardDetailAnswer detail)
        {
            ArgumentNullException.ThrowIfNull(detail);

            IReadOnlyList<BoardHistoryEntry> history = detail.History ?? [];
            Dictionary<long, BoardSubmissionEntry> submissions = detail.Submissions
                .GroupBy(submission => submission.Id)
                .ToDictionary(group => group.Key, group => group.First());

            var items = new List<HistoryItem>();

            foreach (BoardSubmissionEntry submission in submissions.Values)
            {
                List<BoardHistoryEntry> own = history
                    .Where(entry => entry.SubmissionId == submission.Id && CardEvents.Contains(entry.Event))
                    .OrderBy(entry => entry.AtUtc)
                    .ToList();

                items.Add(BoardHistoryDisplay.Card(submission, own));
            }

            foreach (BoardHistoryEntry entry in history)
            {
                if (entry.SubmissionId is long id && submissions.ContainsKey(id) && CardEvents.Contains(entry.Event))
                    continue;

                items.Add(new HistoryEventLine(entry.AtUtc, BoardsDisplay.HistoryWhatRuns(entry), BoardsDisplay.HistoryFooterRuns(entry), BoardsDisplay.HistoryNote(entry)));
            }

            return items
                .Select((item, index) => (item, index))
                .OrderByDescending(pair => pair.item.AtUtc)
                .ThenBy(pair => pair.index)
                .Select(pair => pair.item)
                .GroupBy(item => BoardHistoryDisplay.MonthOf(item.AtUtc))
                .Select(group => new HistoryMonth(group.Key, group.ToList()))
                .ToList();
        }

        // The heading above everything - and, with nothing at all, the whole view.
        public static string Heading(int submissions, int events)
        {
            if (submissions == 0 && events == 0)
                return "Nothing has happened to this board yet";

            var parts = new List<string>();

            if (submissions > 0)
                parts.Add(submissions == 1 ? "1 submission" : $"{submissions.ToString(CultureInfo.InvariantCulture)} submissions");

            if (events > 0)
                parts.Add(events == 1 ? "1 other event" : $"{events.ToString(CultureInfo.InvariantCulture)} other events");

            return $"History - {string.Join(" and ", parts)}, newest first";
        }

        // "October 2026" - a month's heading, in the time zone the dates are shown in.
        public static string MonthOf(DateTimeOffset at) =>
            at.ToLocalTime().ToString("MMMM yyyy", CultureInfo.InvariantCulture);

        // "4 Oct" - the day in the timeline column; the month is the heading above it.
        public static string DayOf(DateTimeOffset at) =>
            at.ToLocalTime().ToString("d MMM", CultureInfo.InvariantCulture);

        // ###########################################################################################
        // One submission's card. Its steps read top to bottom in the order they happened; the state
        // line is where it stands NOW, in CRT's words ("Published to the stable source").
        // ###########################################################################################
        public static HistorySubmissionCard Card(BoardSubmissionEntry submission, IReadOnlyList<BoardHistoryEntry> events)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(events);

            // Each step a line of runs, whoever did it in bold (owner request, 2026-10-09).
            var steps = new List<IReadOnlyList<ReviewNoteRun>>();
            var times = new List<DateTimeOffset> { submission.CreatedUtc };

            BoardHistoryEntry? sent = events.FirstOrDefault(entry => entry.Event == BoardHistoryEvents.Sent);
            string? from = BoardHistoryDisplay.Blank(sent?.Who) ?? BoardHistoryDisplay.Blank(submission.ContactEmail);
            DateTimeOffset sentAt = sent?.AtUtc ?? submission.CreatedUtc;

            if (submission.CreatedUtc != default || sent is not null)
            {
                steps.Add(from is null
                    ? [PersonRuns.Plain($"{SubmissionReceiptPresenter.FormatDate(sentAt)} - sent")]
                    : [PersonRuns.Plain($"{SubmissionReceiptPresenter.FormatDate(sentAt)} - sent by "), PersonRuns.Person(from)]);
            }

            foreach (BoardHistoryEntry entry in events.Where(entry => entry.Event != BoardHistoryEvents.Sent))
            {
                times.Add(entry.AtUtc);

                string date = SubmissionReceiptPresenter.FormatDate(entry.AtUtc);
                string? who = BoardHistoryDisplay.Blank(entry.Who);
                string? detail = BoardHistoryDisplay.Blank(entry.Detail);

                switch (entry.Event)
                {
                    case BoardHistoryEvents.Decided:
                        steps.Add(who is null
                            ? [PersonRuns.Plain($"{date} - {SubmissionReceiptPresenter.DescribeState(detail)}")]
                            : [PersonRuns.Plain($"{date} - {SubmissionReceiptPresenter.DescribeState(detail)} by "), PersonRuns.Person(who)]);
                        break;

                    case BoardHistoryEvents.Amended:
                        steps.Add(detail is null
                            ? [PersonRuns.Plain($"{date} - changed by "), PersonRuns.PersonOr(who, "a maintainer")]
                            : [PersonRuns.Plain($"{date} - changed by "), PersonRuns.PersonOr(who, "a maintainer"), PersonRuns.Plain($" ({detail})")]);
                        break;

                    // The draft discarded is the red mark under the card, not a step - see below.
                    case BoardHistoryEvents.DraftDiscarded:
                        break;
                }
            }

            if (submission.DecidedUtc is DateTimeOffset decided)
                times.Add(decided);

            if (submission.DraftDiscardedUtc is DateTimeOffset discarded)
                times.Add(discarded);

            return new HistorySubmissionCard(
                times.Max(),
                submission.Id,
                $"#{submission.Id.ToString(CultureInfo.InvariantCulture)} - {BoardsDisplay.SubmissionTitle(submission)}",
                SubmissionReceiptPresenter.DescribeState(submission.State),
                steps,
                BoardsDisplay.SubmissionComment(submission),
                submission.DraftDiscardedUtc is DateTimeOffset at ? DraftDiscardWording.Mark(at) : null,
                submission.Changes is SubmissionChanges changes && SubmissionChangeFacts.HasChanges(changes)
                    ? BoardHistoryDisplay.ChangeLines(changes)
                    : [],
                BoardHistoryDisplay.NoSummaryNote(submission));
        }

        // ###########################################################################################
        // Said on a card that went into BETA but carries no summary - published before the server
        // kept them, so there is no board left to compare it with. Null for every other card: one
        // never published has nothing to summarise, and one with a summary shows it.
        // ###########################################################################################
        public static string? NoSummaryNote(BoardSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            if (submission.Changes is not null)
                return null;

            string state = submission.State?.Trim().ToLowerInvariant() ?? string.Empty;

            return state is "merged" or "published"
                ? "What it changed was not recorded - it was published before CRT kept a record of that."
                : null;
        }

        public const string ChangesHeading = "What it changed in BETA:";

        // ###########################################################################################
        // WHAT A SUBMISSION CHANGED, as lines: per sheet a heading, then a line per kind of change
        // under it, the counts in bold ("[2] added: U7, U8"). The names the server recorded follow
        // each count - "and 5 more" when it recorded only the first of them. A new board is
        // counted, not named: its first submission adds every row it has.
        //
        // Files last: added, replaced (other bytes under the same path - a replaced picture that no
        // row's colour would show) and removed, by file name.
        // ###########################################################################################
        public static IReadOnlyList<HistoryChangeLine> ChangeLines(SubmissionChanges changes)
        {
            ArgumentNullException.ThrowIfNull(changes);

            var lines = new List<HistoryChangeLine>();

            if (changes.IsNewBoard)
                lines.Add(new HistoryChangeLine([new ReviewNoteRun("A new board.", IsCount: false)], IsHeading: true));

            foreach (SectionChanges section in changes.Sections ?? [])
            {
                lines.Add(new HistoryChangeLine([new ReviewNoteRun(section.Section, IsCount: false)], IsHeading: true));

                bool named = !changes.IsNewBoard;

                BoardHistoryDisplay.AddKind(lines, section.AddedCount, "added", named ? section.Added : []);
                BoardHistoryDisplay.AddKind(
                    lines,
                    section.ChangedCount,
                    "changed",
                    named
                        ? (section.Changed ?? []).Select(row => (row.Fields ?? []).Count == 0 ? row.Row : $"{row.Row} ({string.Join(", ", row.Fields!)})").ToList()
                        : []);
                BoardHistoryDisplay.AddKind(lines, section.RemovedCount, "removed", named ? section.Removed : []);
                BoardHistoryDisplay.AddKind(
                    lines,
                    section.RenamedCount,
                    "renamed",
                    named ? (section.Renamed ?? []).Select(row => $"{row.From} to {row.To}").ToList() : []);
            }

            FileChanges files = changes.Files ?? FileChanges.None;

            if (files.HasChanges)
            {
                lines.Add(new HistoryChangeLine([new ReviewNoteRun("Files", IsCount: false)], IsHeading: true));

                BoardHistoryDisplay.AddKind(lines, files.AddedCount, "added", BoardHistoryDisplay.FileNames(files.Added));
                BoardHistoryDisplay.AddKind(lines, files.ReplacedCount, "replaced", BoardHistoryDisplay.FileNames(files.Replaced));
                BoardHistoryDisplay.AddKind(lines, files.RemovedCount, "removed", BoardHistoryDisplay.FileNames(files.Removed));
            }

            return lines;
        }

        // "[2] added: U7, U8" - or "[45] added" with nothing named; nothing at all for a zero.
        private static void AddKind(List<HistoryChangeLine> lines, int count, string verb, IReadOnlyList<string>? names)
        {
            if (count <= 0)
                return;

            var runs = new List<ReviewNoteRun>
            {
                new("[", IsCount: false),
                new(count.ToString(CultureInfo.InvariantCulture), IsCount: true),
                new($"] {verb}", IsCount: false)
            };

            List<string> shown = (names ?? []).Where(name => !string.IsNullOrWhiteSpace(name)).ToList();

            if (shown.Count > 0)
            {
                string list = string.Join(", ", shown);
                int more = count - shown.Count;

                runs.Add(new ReviewNoteRun(
                    more > 0 ? $": {list} and {more.ToString(CultureInfo.InvariantCulture)} more" : $": {list}",
                    IsCount: false));
            }

            lines.Add(new HistoryChangeLine(runs, IsHeading: false));
        }

        // A file by its name alone - the folders are the board's own, and the name is what is read.
        private static IReadOnlyList<string> FileNames(IReadOnlyList<string>? paths) =>
            (paths ?? []).Select(path => Path.GetFileName(path.Replace('\\', '/').TrimEnd('/'))).ToList();

        private static string? Blank(string? text) =>
            string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }

    // One month of the History view, newest item first.
    public sealed record HistoryMonth(string Heading, IReadOnlyList<HistoryItem> Items);

    // One item of the History view: a submission's card, or any other event's line.
    public abstract record HistoryItem(DateTimeOffset AtUtc);

    // ###########################################################################################
    // One submission, all on one card. Title: "#14 - what the contributor wrote". State: where it
    // stands now. Steps: what happened to it, oldest first. Told: what the contributor was told.
    // DraftDiscarded: the red mark, when the contributor threw their draft away. Changes: what it
    // changed in BETA, or none. NoSummary: said instead of the changes when it went into BETA before
    // they were recorded.
    // ###########################################################################################
    public sealed record HistorySubmissionCard(
        DateTimeOffset AtUtc,
        long Id,
        string Title,
        string State,
        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> StepRuns,
        string? Told,
        string? DraftDiscarded,
        IReadOnlyList<HistoryChangeLine> Changes,
        string? NoSummary) : HistoryItem(AtUtc)
    {
        // Each step as one text.
        public IReadOnlyList<string> Steps => this.StepRuns.Select(PersonRuns.Text).ToList();
    }

    // Any other event: what happened, the grey line with who and anything more, and - for a decision
    // whose submission has no card - what the contributor was told.
    public sealed record HistoryEventLine(
        DateTimeOffset AtUtc,
        IReadOnlyList<ReviewNoteRun> WhatRuns,
        IReadOnlyList<ReviewNoteRun> FooterRuns,
        string? Note = null) : HistoryItem(AtUtc)
    {
        public string What => PersonRuns.Text(this.WhatRuns);

        public string Footer => PersonRuns.Text(this.FooterRuns);
    }

    // One line of what a submission changed: a sheet's heading, or a kind of change under it.
    public sealed record HistoryChangeLine(IReadOnlyList<ReviewNoteRun> Runs, bool IsHeading)
    {
        public string Text => string.Concat(this.Runs.Select(run => run.Text));
    }
}
