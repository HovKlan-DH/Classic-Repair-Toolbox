using System;
using System.Collections.Generic;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE STAGE LINE under a board's name (owner request, 2026-10-04: "where does this sit now, as I
    // do not think it is in BETA nor stable?") - its newest submission, BETA and the stable source,
    // each with its state. The words are BoardStagesDisplay's; this part only puts them on screen.
    //
    // BETA and the stable source come with the board's row and show at once; the submission comes
    // with its detail, so "Submitted" fills in when that arrives - and only for the board the detail
    // is for.
    // ###########################################################################################
    public partial class BoardDetailView
    {
        // The submissions the stage line is drawn from, and whose board they are - null until the
        // detail of the board on screen arrives.
        private IReadOnlyList<BoardSubmissionEntry>? thisStageSubmissions;
        private string? thisStageSubmissionsFor;

        private static readonly (string Label, string Value, string Detail)[] StageParts =
        [
            ("StageSubmittedLabel", "StageSubmittedValue", "StageSubmittedDetail"),
            ("StageBetaLabel", "StageBetaValue", "StageBetaDetail"),
            ("StageStableLabel", "StageStableValue", "StageStableDetail")
        ];

        // Each stage's card and what the place is (owner request, 2026-10-09).
        private static readonly (string Card, string Note, string Text)[] StageCards =
        [
            ("StageSubmittedCard", "StageSubmittedNote", BoardStagesDisplay.SubmittedNote),
            ("StageBetaCard", "StageBetaNote", BoardStagesDisplay.BetaNote),
            ("StageStableCard", "StageStableNote", BoardStagesDisplay.StableNote)
        ];

        // The detail's submissions, for the board they are of.
        private void UseStageSubmissions(BoardDetailAnswer detail)
        {
            this.thisStageSubmissions = detail.Submissions;
            this.thisStageSubmissionsFor = detail.Board.BoardId;
        }

        private void ShowStages(BoardOverviewEntry? board)
        {
            this.SetShown("BoardStagesPanel", board is not null);
            this.SetShown("BoardStageNowText", false);

            if (board is null)
                return;

            IReadOnlyList<BoardSubmissionEntry>? submissions =
                string.Equals(this.thisStageSubmissionsFor, board.BoardId, StringComparison.Ordinal) ? this.thisStageSubmissions : null;

            IReadOnlyList<BoardStage> stages = BoardStagesDisplay.For(board, submissions);

            for (int i = 0; i < BoardDetailView.StageParts.Length && i < stages.Count; i++)
            {
                BoardStage stage = stages[i];
                (string labelName, string valueName, string detailName) = BoardDetailView.StageParts[i];

                if (this.FindControl<TextBlock>(labelName) is TextBlock label)
                    label.Text = BoardStagesDisplay.NumberedLabel(i, stage.Label);

                if (this.FindControl<TextBlock>(valueName) is TextBlock value)
                {
                    value.Text = stage.Value;
                    value.IsVisible = stage.Value.Length > 0;
                    value.Classes.Set("Quiet", stage.State != BoardStageState.Reached);
                }

                if (this.FindControl<TextBlock>(detailName) is TextBlock detail)
                {
                    detail.Text = stage.Detail ?? string.Empty;
                    detail.IsVisible = !string.IsNullOrEmpty(stage.Detail);
                }
            }

            // Where the newest work sits: its card outlined, and the sentence under the three.
            BoardStageNow? now = BoardStagesDisplay.Now(board, submissions);

            for (int i = 0; i < BoardDetailView.StageCards.Length; i++)
            {
                (string cardName, string noteName, string note) = BoardDetailView.StageCards[i];

                if (this.FindControl<Border>(cardName) is Border card)
                    card.Classes.Set("Current", now?.Stage == i);

                if (this.FindControl<TextBlock>(noteName) is TextBlock noteText)
                    noteText.Text = note;
            }

            if (now is not null && this.FindControl<TextBlock>("BoardStageNowText") is TextBlock nowText)
            {
                TabMaintainer.ShowCounts(nowText, [new ReviewNoteRun("Now: ", IsCount: true), new ReviewNoteRun(now.Sentence, IsCount: false)]);
                nowText.IsVisible = true;
            }
        }

        // Which stage is drawn as current (0 to 2), or -1 - and the sentence under them - for tests.
        internal int CurrentStageForTests =>
            Array.FindIndex(BoardDetailView.StageCards, stage => this.FindControl<Border>(stage.Card)?.Classes.Contains("Current") == true);

        internal string StageNowForTests =>
            this.FindControl<TextBlock>("BoardStageNowText") is { IsVisible: true } now ? TabMaintainer.TextOf(now) : string.Empty;

        // The stage line as shown, one "Label: value - detail" per stage - for tests.
        internal IReadOnlyList<string> StagesForTests()
        {
            var lines = new List<string>();

            foreach ((string labelName, string valueName, string detailName) in BoardDetailView.StageParts)
            {
                string label = this.FindControl<TextBlock>(labelName)?.Text ?? string.Empty;
                TextBlock? value = this.FindControl<TextBlock>(valueName);
                TextBlock? detail = this.FindControl<TextBlock>(detailName);

                string line = $"{label}: {(value is { IsVisible: true } ? value.Text : string.Empty)}";

                if (detail is { IsVisible: true })
                    line += $" - {detail.Text}";

                lines.Add(line);
            }

            return lines;
        }

        internal bool StagesShownForTests => this.FindControl<Grid>("BoardStagesPanel")?.IsVisible == true;
    }
}
