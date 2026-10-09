using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // A BOARD'S SIX VIEWS ON THE "BOARDS" SCREEN (owner request, 2026-10-03: "I would like all
    // the same functionalities, as the 'Contributor Submissions' has. I would like to see these
    // options when selecting a system: Board data (and I should be able to do the same edits),
    // Files (should not show changed files - just list all files), Contributor, Maintainer,
    // Statistics ... where 'Maintainer' is the info as it has now").
    //
    //   Board data  - BETA's board in the submission's own table, editable by the board's
    //                 maintainers and the administrator. A save goes STRAIGHT TO BETA (owner
    //                 decision, 2026-10-03: "it should go directly to the next queue, 'BETA > Stable',
    //                 so it can directly be tested in BETA"), with a reason asked for when Save is
    //                 pressed - not while the board waits under BETA > Stable.
    //   Files       - every file the board uses as BETA holds it, nothing marked as changing.
    //   Contributor - who has contributed to it, and how that went.
    //   Maintainer  - its place in CRT's drop-down lists while it needs one, and who maintains it -
    //                 a list and nothing more (owner request, 2026-10-04: "The stuff that should be
    //                 visible in here, is just the selected maintainer(s)"). Adding, inviting and
    //                 removing them is the administrator's, under Account > Maintainers.
    //   History     - everything that has happened to it, one card per submission with what it
    //                 changed (owner request, 2026-10-04 - BoardHistoryDisplay).
    //   Statistics  - how often CRT users look at it (the board views). Graphs come later.
    //
    // Board data first, like a submission's own views (SubmissionViews) - a row of buttons reads
    // as opening on its first. But a board is not a submission: the views are about "a board",
    // so the one chosen STAYS when another board is chosen - looking at each board's statistics
    // or maintainers in turn is the ordinary use, and starting over on Board data every time would
    // cost a click and a table read per board.
    //
    // Pure, so the order, the labels and the words are tested.
    // ###########################################################################################
    public enum BoardSection
    {
        BoardData,
        Files,
        Contributor,
        Maintainer,
        History,
        Statistics
    }

    public static class BoardSections
    {
        // The order of the buttons, left to right.
        public static IReadOnlyList<BoardSection> Order { get; } =
        [
            BoardSection.BoardData,
            BoardSection.Files,
            BoardSection.Contributor,
            BoardSection.Maintainer,
            BoardSection.History,
            BoardSection.Statistics
        ];

        // What the screen opens on before any view is chosen.
        public const BoardSection Opening = BoardSection.BoardData;

        public static string Label(BoardSection section) => section switch
        {
            BoardSection.Files => "Files",
            BoardSection.Contributor => "Contributor",
            BoardSection.Maintainer => "Maintainer",
            BoardSection.History => "History",
            BoardSection.Statistics => "Statistics",
            _ => SubmissionViews.Label(SubmissionView.BoardData)
        };

        // ###########################################################################################
        // WHY BETA'S TABLE CANNOT BE CHANGED - the amber panel above the views (owner request,
        // 2026-10-09), said whichever view is open. The server's reason - from the table, or from the
        // board's detail before the table is read (code review, 2026-10-09) - or null when it can be.
        // ###########################################################################################
        public static string? ReadOnlyReason(bool mayEdit, string? mayNotEditReason)
        {
            if (mayEdit)
                return null;

            return string.IsNullOrWhiteSpace(mayNotEditReason)
                ? "You can look at this board's data, but not send a change to it."
                : mayNotEditReason.Trim();
        }

        public static string? ReadOnlyReason(BoardTableAnswer table)
        {
            ArgumentNullException.ThrowIfNull(table);
            return BoardSections.ReadOnlyReason(table.MayEdit, table.MayNotEditReason);
        }

        // ###########################################################################################
        // The line above BETA's table, under the BETA / Stable switch: `said` first - what a publish
        // just did, say - then what the table is. One that can be changed says what it is coloured
        // against and what a save does; one that cannot says only what it is compared with, when
        // compared (code review, 2026-10-09: its colours were otherwise explained nowhere) - why it
        // cannot be changed is the panel's (ReadOnlyReason). Null when there is nothing to say.
        // ###########################################################################################
        public static string? TableNote(BoardTableAnswer table, bool comparedWithStable, string? said = null)
        {
            ArgumentNullException.ThrowIfNull(table);

            string? what;

            if (table.MayEdit)
            {
                string marked = comparedWithStable
                    ? "BETA's board compared with the stable source - everything that differs from it is marked, and so is what you change."
                    : "BETA's board as it is now - only what you change is marked.";

                what = marked + " Saving asks for a reason and publishes your change straight to BETA, where it waits under " +
                       MaintainerScreenWording.BetaQueueQuoted + " for the stable source.";
            }
            else
            {
                what = comparedWithStable ? BoardSections.ComparedReadOnlyLine : null;
            }

            string note = string.Join(" ", new[] { said?.Trim(), what }.Where(part => !string.IsNullOrWhiteSpace(part)));

            return note.Length == 0 ? null : note;
        }

        public const string ComparedReadOnlyLine =
            "BETA's board compared with the stable source - everything that differs from it is marked.";

        // ###########################################################################################
        // "COMPARE SOURCES" (owner request, 2026-10-09: "a 'Compare sources' checkbox shown right after
        // the two radio buttons ... the below table should show the changes as-if this was a normal
        // board submission"). Ticked, BETA's table is coloured against the stable source and the
        // stable source's against BETA, and a changed cell's tooltip names the OTHER source's value -
        // "Stable source value:" or "BETA source value:", the Drafts tab's words
        // (BoardTableDocument.SourceBaselineLabel). Remembered between launches.
        //
        // Only a board in BOTH can be compared, and never while BETA's table holds a change not
        // saved: comparing opens the table again, which would throw the change away.
        // ###########################################################################################
        public const string CompareSourcesLabel = "Compare sources";

        public const string CompareSourcesTip =
            "Mark everything that differs between BETA and the stable source - a changed cell gives the other source's value.";

        public const string CompareNeedsBothSources = "Only a board in both BETA and the stable source can be compared.";

        public const string CompareNeedsSavedTable = "Save or undo your change in BETA's table first - comparing opens the table again.";

        public static bool CanCompareSources(BoardOverviewEntry? board) => board is { InBeta: true, InProduction: true };

        // Why the box cannot be used now, or null when it can.
        public static string? CompareUnavailable(BoardOverviewEntry? board, bool betaHoldsUnsavedChange)
        {
            if (!BoardSections.CanCompareSources(board))
                return BoardSections.CompareNeedsBothSources;

            return betaHoldsUnsavedChange ? BoardSections.CompareNeedsSavedTable : null;
        }

        // A file cell's card, compared: each picture headed by the source it is from.
        public const string StableSourceSide = "Stable source";

        public const string BetaSourceSide = "BETA source";

        // ###########################################################################################
        // What the table calls its two sides - a changed cell's tooltip names the value it replaced
        // by the first, and a file cell's card heads its two pictures with both. "BETA now", since
        // the table is coloured against BETA as it was opened (there is no "published" in between).
        // ###########################################################################################
        public const string BaselineLabel = "BETA now";

        public const string ChangeLabel = "Your change";

        // The two "please wait"s of a save: the check of what it would remove, then the publish -
        // which the server checks as it would any submission.
        public const string CheckingWait = "Checking your change against BETA...";

        public const string PublishingWait =
            "Publishing your change to BETA. The server checks it as it would any submission - please wait until it is done.";

        // ###########################################################################################
        // THE REASON, ASKED WHEN SAVE IS PRESSED (owner request, 2026-10-03: "do ask for a change
        // reason when clicking the 'Save changes' button, so this can go along with the change, just
        // like any normal contribution") - PublishBoardChangeWindow's words. It replaced a description
        // box above the table. With the files the publish would remove, when there are any: the
        // maintainer sees them before anything is written, as before an approval.
        // ###########################################################################################
        public const string ReasonTitle = "Publish your change to BETA";

        public static string ReasonHeadline(string boardId) => $"Publish your change to {boardId} in BETA";

        // The BETA check box named as the Configuration tab names it - one constant, the mails' too.
        public static string ReasonExplanation { get; } =
            "Your change goes straight into the BETA data, where you can try it in CRT with \"" +
            ConfigurationWording.BetaSourceCheckBox + "\" ticked in the Configuration tab. " +
            "It then waits under " + MaintainerScreenWording.BetaQueueQuoted + " until it is published to the stable source - and until then no other " +
            "change can be made to this board.";

        public const string ReasonLabel = "Reason for the change";

        public const string ReasonHint = "Corrected U8's part number.";

        public const string ReasonNote =
            "A sentence is enough. It goes with the change, like a contribution's description - under " + MaintainerScreenWording.BetaQueueQuoted + " and in the board's history.";

        public const string PublishButton = "Publish to BETA";

        // Null when the change removes nothing - the dialog then shows no list at all.
        public static string? RemovalsHeading(IReadOnlyCollection<string> removals)
        {
            ArgumentNullException.ThrowIfNull(removals);

            return removals.Count switch
            {
                0 => null,
                1 => "This also removes 1 file from BETA - nothing uses it once your change is in:",
                _ => $"This also removes {removals.Count} files from BETA - nothing uses them once your change is in:"
            };
        }

        // ###########################################################################################
        // After the change was published: at which revision, what it removed, and any warnings the
        // server raised about it. The table then shows BETA as it is now - read-only, with the
        // server's reason, while the board waits under BETA > Stable.
        // ###########################################################################################
        public static string Published(BoardEditResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            string published = string.IsNullOrWhiteSpace(result.Revision)
                ? "Published to BETA."
                : $"Published to BETA as revision {result.Revision.Trim()}.";

            int removed = result.RemovedFiles?.Count ?? 0;

            if (removed > 0)
                published += removed == 1 ? " 1 file nothing used any more was removed." : $" {removed} files nothing used any more were removed.";

            return BoardSections.WithWarnings(published, result);
        }

        // ###########################################################################################
        // Made into a submission but NOT published - something changed between the check and the
        // publish, say. Nothing is lost: it waits under Contributor Submissions, where the screen then
        // opens it.
        // ###########################################################################################
        public static string SavedNotPublished(BoardEditResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            string reason = string.IsNullOrWhiteSpace(result.NotPublishedReason)
                ? "the server did not say why."
                : result.NotPublishedReason.Trim();

            return BoardSections.WithWarnings(
                $"Saved as submission #{result.SubmissionId}, but not published to BETA: {reason} " +
                "It waits under " + MaintainerScreenWording.ContributorQueueQuoted + ", where you can approve it.",
                result);
        }

        private static string WithWarnings(string line, BoardEditResult result)
        {
            List<string> warnings = result.Warnings
                .Select(warning => warning.Message)
                .Where(message => !string.IsNullOrWhiteSpace(message))
                .ToList();

            return warnings.Count == 0 ? line : $"{line} Warnings: {string.Join(" ", warnings)}";
        }

        // After a publish whose table could not be read again: the table is closed rather than left
        // on BETA as it was before the change.
        public static string NotReadAgain(string reason) =>
            $"The table could not be read again ({(string.IsNullOrWhiteSpace(reason) ? "no answer" : reason.Trim().TrimEnd('.'))}) - " +
            "choose another view and come back to Board data to read it.";

        // ###########################################################################################
        // BETA's board moved while the table holds a change not published yet (2026-10-04) - the
        // table is then not read again under it, and the server would refuse the change
        // (BoardEditFlow.ChangedSinceMessage), so it is said before it is tried. Undone, the
        // change is gone and Board data reads BETA again when next shown.
        // ###########################################################################################
        public const string BetaMovedUnderChange =
            "This board's data in BETA has changed since you opened the table, so your change can no longer be published. " +
            "Copy anything you need, undo your changes (Ctrl+Z), then choose another view and come back to Board data to see BETA as it is now.";

        public static string NotPublished(string reason) =>
            $"Not published: {(string.IsNullOrWhiteSpace(reason) ? "the server refused it." : reason)}";

        // ###########################################################################################
        // AFTER NO ANSWER IN TWO MINUTES: the submission a change became, if it arrived - the board's
        // newest submission with this reason as its description. Its state then says how far it got:
        // in BETA ("merged", or "published" once it went on to stable), or waiting in the queue.
        // ###########################################################################################
        public static BoardSubmissionEntry? FindSent(BoardDetailAnswer detail, string reason)
        {
            ArgumentNullException.ThrowIfNull(detail);

            string wanted = reason?.Trim() ?? string.Empty;

            return detail.Submissions
                .Where(submission => string.Equals(submission.Summary?.Trim(), wanted, StringComparison.Ordinal))
                .OrderByDescending(submission => submission.Id)
                .FirstOrDefault();
        }

        public static bool ReachedBeta(BoardSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.Equals(submission.State, "merged", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(submission.State, "published", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // WHAT A SAVE SENDS: BETA's rows with the table's edits applied, and the KiCad calibrations as
        // BETA has them (the table cannot show them). The server keeps highlights, calibrations and
        // the revision date as BETA has them whatever is sent (SubmissionRowsBoard.WithTableSections)
        // - the same rule a submission's own table saves by (ReviewTableWording.RowsToSave).
        // ###########################################################################################
        public static SubmissionRows RowsToSend(BoardTableDocument document, SubmissionRows beta) =>
            ReviewTableWording.RowsToSave(document, beta);

        // The Files view's line above the tree.
        public const string FilesHeading =
            "Every file this board uses in BETA: its own folder, and the shared files its board cites.";

        // ###########################################################################################
        // THE BETA / STABLE SWITCH ON BOARD DATA AND FILES (owner request, 2026-10-04: "Shouldn't there
        // be somewhere a possibility to see what we actually do have in BETA or stable for this?").
        // BETA's table is the one a change is made in; the stable source's is read-only - it changes
        // only when BETA is published to it.
        //
        // Which is SHOWN: the one wanted when the board is there; otherwise the stable source when
        // only that holds it; otherwise BETA, whose view then says nothing is there. The one wanted
        // stays wanted for the next board, as the chosen view does.
        // ###########################################################################################
        public static bool ShowsStable(bool stableWanted, BoardOverviewEntry? board)
        {
            if (board is null || board.InProduction != true)
                return false;

            return stableWanted || !board.InBeta;
        }

        // Whether each half of the switch can be chosen: the stable source only when it holds the
        // board, BETA unless only the stable source does - BETA's view says so when neither does.
        public static bool CanChooseStable(BoardOverviewEntry? board) => board?.InProduction == true;

        public static bool CanChooseBeta(BoardOverviewEntry? board) =>
            board is not null && (board.InBeta || board.InProduction != true);

        // The stable source's table: what it is called beside a changed cell (it never has one), and
        // in "there is no file at this path in ...".
        public const string StableBaselineLabel = "Stable now";

        public const string StableTreeName = "the stable source";

        // And BETA, where a file cell's file is not: "There is no file at this path in BETA."
        public const string BetaTreeName = "BETA";

        // ###########################################################################################
        // *** AN ANSWER THAT IS BETA'S IS NOT SHOWN AS STABLE. *** A server older than 4.5.0 does not
        // know the request's `tree` and answers with BETA's board and files - which, drawn under
        // "Stable", would tell the maintainer BETA is what everybody has. The stable source's answer
        // names no BETA address, and its files open from the stable source; anything else is BETA's.
        // ###########################################################################################
        public static bool IsStableAnswer(BoardTableAnswer table)
        {
            ArgumentNullException.ThrowIfNull(table);
            return table.BetaDataUrl is null && !table.MayEdit;
        }

        public static bool IsStableAnswer(BoardFilesAnswer files)
        {
            ArgumentNullException.ThrowIfNull(files);
            return files.BetaDataUrl is null && files.Files.All(file => file.OpenFrom == BoardFileSource.Production);
        }

        public const string StableNeedsNewerServer =
            "This server cannot show the stable source yet - it needs CRT.Server 4.5.0 or newer.";

        public const string StableFilesHeading =
            "Every file this board uses in the stable source - what everybody using CRT gets: its own folder, and the shared files its board cites.";

        // ###########################################################################################
        // Above the views, for a board not in CRT's drop-down lists: the list's own mark - and, while
        // it still NEEDS a place, where that is given: the Maintainer view. A view chosen for another
        // board stays chosen, so the placement could otherwise go unseen. Null for every other board.
        // ###########################################################################################
        public static string? PlacementLine(BoardListingAnswer? listing, string? boardId)
        {
            if (BoardPlacementDisplay.ListMark(listing, boardId) is not string mark)
                return null;

            return BoardPlacementDisplay.UnlistedEntry(listing, boardId)?.Placement is null
                ? $"{mark} - give it one under Maintainer"
                : mark;
        }
    }
}
