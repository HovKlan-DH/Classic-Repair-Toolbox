using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE PROBLEM IN A BOARD'S DATA, AND WHERE IT IS (owner request, 2026-10-02: show the checks
    // in the Drafts tab's table "as either Error or Warning ... All error should be fixed before
    // submission can be done").
    //
    // Found by BoardDataChecks - the one set of rules the table, the launch log (DataValidator)
    // and the server's SubmissionValidator all use. Where it is, is what the table needs and the
    // server's own findings never carried: the SHEET, the row's INDEX in that sheet's section of
    // BoardData (the entry order, which is the table's live, non-blank rows in order), and the
    // COLUMN. A problem that is in no sheet - a highlight, which lives in the JSON beside the
    // workbook - has no sheet and index -1.
    //
    // Code, Subject and Message are exactly what the server's ValidationFinding carries for the
    // same problem, so the table and a refusal from the server say the same words.
    // ###########################################################################################
    public enum BoardProblemLevel
    {
        None = 0,

        // Worth a look, never in the way: the data still loads and the server still accepts it.
        Warning = 1,

        // The server refuses a submission carrying it (or the client refuses to send it), so it
        // has to be fixed before submitting.
        Error = 2
    }

    // How many errors and warnings a board has - what a draft's row on the Drafts tab shows
    // (DraftStatusReader.CountProblemsCached).
    public readonly record struct BoardProblemCounts(int Errors, int Warnings)
    {
        public static BoardProblemCounts None { get; } = new(0, 0);

        public static BoardProblemCounts Of(IEnumerable<BoardDataProblem> problems)
        {
            ArgumentNullException.ThrowIfNull(problems);

            List<BoardDataProblem> all = problems.ToList();

            return new BoardProblemCounts(
                all.Count(problem => problem.Level == BoardProblemLevel.Error),
                all.Count(problem => problem.Level == BoardProblemLevel.Warning));
        }
    }

    public sealed record BoardDataProblem(
        BoardProblemLevel Level,
        string Code,
        string Subject,
        string Message,
        string? Sheet,
        int Index,
        string? Column)
    {
        public bool IsInSheet => this.Sheet is not null && this.Index >= 0;
    }

    // ###########################################################################################
    // Which rules to run.
    //
    //   Submission - exactly the server's row rules, with its codes, subjects and messages, in its
    //                order. SubmissionValidator runs these; a new rule here would change what the
    //                server refuses, so nothing is added to this scope lightly.
    //   Everything - those, plus the file-name rules the server applies elsewhere
    //                (SubmissionPathRules.IsSafelyShaped, SubmissionFileRules.TryCheckName) and the
    //                local warnings DataValidator used to log on its own (oscilloscope settings,
    //                orphan rows, components with no highlight, mixed part numbers). Every one of
    //                the added rules is a WARNING, except those file-name rules, which the server
    //                refuses - so an error in the table is always a refusal at submit.
    // ###########################################################################################
    public enum BoardCheckScope
    {
        Submission,
        Everything
    }
}
