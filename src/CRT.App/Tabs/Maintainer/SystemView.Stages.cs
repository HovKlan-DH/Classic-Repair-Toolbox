using System;
using System.Collections.Generic;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE STAGE LINE under a system's name (owner request, 2026-10-04: "where does this sit now, as I
    // do not think it is in BETA nor stable?") - its newest submission, BETA and the stable source,
    // each with its state. The words are SystemStagesDisplay's; this part only puts them on screen.
    //
    // BETA and the stable source come with the system's row and show at once; the submission comes
    // with its detail, so "Submitted" fills in when that arrives - and only for the system the detail
    // is for.
    // ###########################################################################################
    public partial class SystemView
    {
        // The submissions the stage line is drawn from, and whose system they are - null until the
        // detail of the system on screen arrives.
        private IReadOnlyList<SystemSubmissionEntry>? thisStageSubmissions;
        private string? thisStageSubmissionsFor;

        private static readonly (string Label, string Value, string Detail)[] StageParts =
        [
            ("StageSubmittedLabel", "StageSubmittedValue", "StageSubmittedDetail"),
            ("StageBetaLabel", "StageBetaValue", "StageBetaDetail"),
            ("StageStableLabel", "StageStableValue", "StageStableDetail")
        ];

        // The detail's submissions, for the system they are of.
        private void UseStageSubmissions(SystemDetailAnswer detail)
        {
            this.thisStageSubmissions = detail.Submissions;
            this.thisStageSubmissionsFor = detail.System.SystemId;
        }

        private void ShowStages(SystemOverviewEntry? system)
        {
            this.SetShown("SystemStagesPanel", system is not null);

            if (system is null)
                return;

            IReadOnlyList<SystemSubmissionEntry>? submissions =
                string.Equals(this.thisStageSubmissionsFor, system.SystemId, StringComparison.Ordinal) ? this.thisStageSubmissions : null;

            IReadOnlyList<SystemStage> stages = SystemStagesDisplay.For(system, submissions);

            for (int i = 0; i < SystemView.StageParts.Length && i < stages.Count; i++)
            {
                SystemStage stage = stages[i];
                (string labelName, string valueName, string detailName) = SystemView.StageParts[i];

                if (this.FindControl<TextBlock>(labelName) is TextBlock label)
                    label.Text = stage.Label;

                if (this.FindControl<TextBlock>(valueName) is TextBlock value)
                {
                    value.Text = stage.Value;
                    value.IsVisible = stage.Value.Length > 0;
                    value.Classes.Set("Quiet", stage.State != SystemStageState.Reached);
                }

                if (this.FindControl<TextBlock>(detailName) is TextBlock detail)
                {
                    detail.Text = stage.Detail ?? string.Empty;
                    detail.IsVisible = !string.IsNullOrEmpty(stage.Detail);
                }
            }
        }

        // The stage line as shown, one "Label: value - detail" per stage - for tests.
        internal IReadOnlyList<string> StagesForTests()
        {
            var lines = new List<string>();

            foreach ((string labelName, string valueName, string detailName) in SystemView.StageParts)
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

        internal bool StagesShownForTests => this.FindControl<Grid>("SystemStagesPanel")?.IsVisible == true;
    }
}
