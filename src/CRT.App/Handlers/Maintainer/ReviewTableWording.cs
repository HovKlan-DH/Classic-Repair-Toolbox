using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the maintainer's table reads (owner request, 2026-09-25: "make the same 'Edit in table
    // format' (maybe call it 'View in table format') available in the Maintainer tab ... The maintainer
    // should be able to also edit whatever, if he chooses to publish it afterwards").
    //
    // It opens with the submission since 2026-09-26 - there is no button for it any more.
    //
    // Pure, so the words - and the one decision here, what a save sends - are tested.
    // ###########################################################################################
    public static class ReviewTableWording
    {
        // ###########################################################################################
        // Why a decision waits while the table holds unsaved changes. A decision is about what the
        // SERVER holds, and approving with changes still on screen would publish the version
        // without them - the opposite of what the maintainer is looking at.
        // ###########################################################################################
        public const string SaveTableBeforeDeciding =
            "The table has changes that are not saved. Save them (or undo them) first - a decision is about the submission as it is saved.";

        // What the table is coloured against, under the toolbar when it opens.
        public static string OpenedMessage(ReviewTableData table)
        {
            ArgumentNullException.ThrowIfNull(table);

            return table.Published is null
                ? "Nothing of this system is published yet, so only a change you make here is marked. A change you save becomes the submission's content before you decide on it."
                : "Coloured against the published board. A change you save here becomes the submission's content before you decide on it.";
        }

        // After a save: what it did, and any warnings the server raised about the new content.
        public static string Saved(ReviewAmendResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            string saved = "Saved - the submission now reads as shown. Any approval given before this change was cleared.";

            List<string> warnings = result.Warnings
                .Select(warning => warning.Message)
                .Where(message => !string.IsNullOrWhiteSpace(message))
                .ToList();

            return warnings.Count == 0
                ? saved
                : $"{saved} Warnings: {string.Join(" ", warnings)}";
        }

        // The "please wait" over the whole window while a save is on its way (2026-09-28) - the
        // server checks the new content as it would a new submission, which takes a moment on a
        // large board. See ReviewDecisionWording.Waiting.
        public const string SavingWait =
            "Saving your changes to the submission. The server checks them as it would a new submission - please wait until it is done.";

        public static string NotSaved(string reason) =>
            $"Not saved: {(string.IsNullOrWhiteSpace(reason) ? "the server refused it." : reason)}";

        // ###########################################################################################
        // The line on the submission view once a maintainer has changed it. Null when nobody has.
        // ###########################################################################################
        public static string? AmendedLine(ReviewAmendmentView? amendment)
        {
            if (amendment is null)
                return null;

            string when = amendment.AtUtc is DateTimeOffset at
                ? " on " + at.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)
                : string.Empty;

            string by = string.IsNullOrWhiteSpace(amendment.By) ? "a maintainer" : amendment.By;

            return $"Changed in the table by {by}{when}. What is shown is the changed version; the contributor is told it was changed.";
        }

        // ###########################################################################################
        // WHAT A SAVE SENDS: the submission's rows with the table's edits applied. The table covers
        // nine sheets; everything else (highlights, calibrations, the revision date) is left as it
        // is, and the server takes those from the submission regardless - moving highlights and
        // calibrations with a schematic the table deleted or renamed
        // (SubmissionRowsBoard.WithTableSections).
        // ###########################################################################################
        public static SubmissionRows RowsToSave(BoardTableDocument document, SubmissionRows submitted)
        {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(submitted);

            SubmissionRows rows = SubmissionRowsBoard.FromBoard(document.ApplyTo(SubmissionRowsBoard.ToBoard(submitted)));
            rows.KiCadCalibrations = submitted.KiCadCalibrations.ToList();

            return rows;
        }
    }
}
