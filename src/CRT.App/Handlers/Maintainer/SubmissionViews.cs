using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // THE THREE VIEWS OF A SUBMISSION (owner request, 2026-09-30: "The "Files..." button ... can
    // easily be missed, and then this great possibility may not be used ... Maybe one row with
    // "Contributor info", "Board data", and then "Files"?").
    //
    //   Board data  - the table, as it always was: the nine sheets, and above them what the table
    //                 cannot show (ReviewNotInTable). The one a submission OPENS on - the owner's
    //                 "make the table the default first view" (2026-09-26) still stands.
    //   Files       - the board's files after approving, against BETA now (FileTreeView). Was the
    //                 "Files..." button's window.
    //   Contributor - who sent it and everything they have sent before (ReviewContributorHistory).
    //
    // Board data comes FIRST, not in the order the owner floated: the view a row of buttons opens
    // on reads as the first of them, and a row whose chosen button sits in the middle looks as if
    // something had already been clicked.
    //
    // *** THE FILES BUTTON CARRIES A COUNT, AND THAT COUNT IS NOT DECORATION. *** It took over from
    // the lines above the table that said "Files: [3] included ([1] replaced under the same name)"
    // and "KiCad data included: ..." - lines that existed because a file replaced under its own
    // path colours nothing in the table. With them gone, the count is what still says so before
    // anything is opened. Pure, so it is tested - and tested to agree with the tree.
    // ###########################################################################################
    public enum SubmissionView
    {
        BoardData,
        Files,
        Contributor
    }

    public static class SubmissionViews
    {
        // The order of the buttons, left to right.
        public static IReadOnlyList<SubmissionView> Order { get; } =
            [SubmissionView.BoardData, SubmissionView.Files, SubmissionView.Contributor];

        // What a newly chosen submission opens on.
        public const SubmissionView Opening = SubmissionView.BoardData;

        public static string Label(SubmissionView view) => view switch
        {
            SubmissionView.Files => "Files",
            SubmissionView.Contributor => "Contributor",
            _ => "Board data"
        };

        // ###########################################################################################
        // The files the SUBMISSION adds, replaces or removes - the Files button's count.
        //
        // *** THROUGH THE TREE'S OWN RULE. *** BoardFileEntries.ForApproval is what the server
        // builds the tree with, and what the tree then counts; asking it here, with the same facts
        // the submission detail already carries, is what keeps the button and the tree's headline
        // from ever disagreeing about a file. Left out, as the headline leaves them out: the
        // workbook and the highlight file, which the approval writes from the table (the Board data
        // view) and which the detail does not name.
        // ###########################################################################################
        public static int ChangingFiles(IReadOnlyList<SubmittedFileFact>? submitted, FileRemovalPreview? removals) =>
            BoardFileEntries
                .ForApproval(submitted, removals?.Files, ownInBeta: null, workbook: null, sidecar: null)
                .Count(entry => entry.Change != BoardFileChange.Unchanged);

        // The button's count, or null for no badge: a submission that changes no file has nothing
        // to look at there, and a "0" would read as something to check.
        public static string? FilesBadge(int changing) =>
            changing <= 0 ? null : changing.ToString(CultureInfo.InvariantCulture);
    }
}
