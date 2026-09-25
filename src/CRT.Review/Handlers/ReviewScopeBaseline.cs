using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT.Review.Handlers
{
    // ###########################################################################################
    // The OSCILLOSCOPE half of task 4's visual review (NewContributeStrategy.md Phase 5).
    //
    // *** THE PICTURE ALONE IS NOT THE REVIEW, AND THAT IS WHY THIS EXISTS. *** A scope baseline
    // is stored as an ordinary component image, so the side-by-side comparison already shows the
    // waveform. What it cannot show is that the SETTINGS moved underneath it - and a trace
    // captured at 2 V/div next to one at 5 V/div looks like a different signal when it is the same
    // signal drawn at a different scale.
    //
    // That is the commonest way a baseline goes wrong: somebody recaptures on a scope configured
    // differently and does not mention it. To a reviewer comparing two pictures it reads as a
    // genuine change in the circuit, and approving it replaces a good reference with one that
    // cannot be compared against anything.
    //
    // So this finds the setting changes and the panel says plainly that the two traces are not
    // directly comparable. It is deliberately a note ATTACHED to the image comparison rather than
    // a separate screen - the pictures are what the reviewer is looking at, and the warning is
    // only useful beside them.
    //
    // *** WHAT THIS IS NOT. *** It does not re-plot the waveform. The baseline is a captured
    // IMAGE, not sampled data - CRT stores what the scope drew, not the points behind it - so
    // there is nothing to re-plot at a common scale even in principle. Saying the scales differ is
    // the honest answer and the whole of what can be said.
    // ###########################################################################################
    public static class ReviewScopeBaseline
    {
        // ###########################################################################################
        // The section and the three columns, read from CRT.Data rather than spelled out.
        //
        // These match against a field diff the SERVER built. A literal that drifts means the
        // setting is never reported and nothing indicates it was missed - the screen looks correct
        // and quietly omits the warning it exists for.
        // ###########################################################################################
        public const string SectionName =
            global::Handlers.DataHandling.BoardWorkbookSchema.SheetComponentImages;

        public const string FieldTimeDiv =
            global::Handlers.DataHandling.BoardWorkbookSchema.ColTimeDiv;

        public const string FieldVoltsDiv =
            global::Handlers.DataHandling.BoardWorkbookSchema.ColVoltsDiv;

        public const string FieldTriggerLevel =
            global::Handlers.DataHandling.BoardWorkbookSchema.ColTriggerLevelVolts;

        // ###########################################################################################
        // Which of the three scope settings moved on one component-image row.
        //
        // Returns false when none did. A component image row changes for all sorts of ordinary
        // reasons - a corrected note, a renamed capture - and reporting those as a settings change
        // would put a warning on a submission that never touched the scope. A warning that fires
        // on everything is one a reviewer learns to ignore, which costs more than it gains.
        // ###########################################################################################
        public static bool TryReadChange(
            ReviewSectionView? section,
            string rowKey,
            out ReviewScopeSettingsChange change)
        {
            change = default;

            if (section is null || string.IsNullOrWhiteSpace(rowKey))
                return false;

            if (!section.FieldChanges.TryGetValue(rowKey, out IReadOnlyList<ReviewFieldChangeView>? fields))
                return false;

            var settings = new List<ReviewScopeSetting>();

            // In the order a scope's own controls read, not the order the diff happens to list
            // them, so two reviewers see the same list in the same order.
            foreach (string name in new[]
            {
                ReviewScopeBaseline.FieldTimeDiv,
                ReviewScopeBaseline.FieldVoltsDiv,
                ReviewScopeBaseline.FieldTriggerLevel
            })
            {
                ReviewFieldChangeView? field = fields.FirstOrDefault(candidate =>
                    string.Equals(candidate.Field, name, StringComparison.Ordinal));

                if (field is not null)
                    settings.Add(new ReviewScopeSetting(field.Field, field.Before, field.After));
            }

            if (settings.Count == 0)
                return false;

            change = new ReviewScopeSettingsChange(settings);
            return true;
        }

        // ###########################################################################################
        // The warning, as the reviewer reads it.
        //
        // *** IT NAMES THE CONSEQUENCE, NOT JUST THE FACT. *** "Settings changed" is true and
        // leaves the reader to work out what it means for the two pictures above it. "...so these
        // two traces are not drawn at the same scale" is the thing they actually need in order to
        // judge what they are looking at.
        //
        // A CLEARED setting reads as "(blank)", the same rule the field-level diff follows - an
        // empty right-hand side reads as the app having failed to draw something, and clearing the
        // settings leaves a baseline nobody can reproduce.
        // ###########################################################################################
        public static string Describe(ReviewScopeSettingsChange change)
        {
            if (change.Settings is null || change.Settings.Count == 0)
            {
                throw new ArgumentException(
                    "A scope settings change must carry at least one setting.", nameof(change));
            }

            string moved = string.Join(", ", change.Settings.Select(setting =>
                $"{setting.Field} {ReviewScopeBaseline.Value(setting.Before)} -> {ReviewScopeBaseline.Value(setting.After)}"));

            return $"Scope settings changed ({moved}), so these two traces are not drawn at the same scale.";
        }

        private static string Value(string raw) => string.IsNullOrEmpty(raw) ? "(blank)" : raw;

        // ###########################################################################################
        // A row key as a HUMAN reads it.
        //
        // *** THE RAW KEY IS UNREADABLE ON SCREEN. *** BoardDraftNaturalKeys joins its parts with
        // U+241F, which renders as a box or as nothing at all depending on the font - so
        // "U8␟ASSY 250407␟12␟Clock" reaches the reviewer as "U8ASSY 250407 12Clock" or worse. The
        // parts are real and useful (which component, which region, which pin), so they are
        // separated visibly rather than hidden.
        //
        // Empty parts are dropped: a component image with no region carries an empty segment, and
        // " / / 12" reads as missing data rather than as a field that does not apply.
        // ###########################################################################################
        public static string DescribeRowKey(string? rowKey)
        {
            if (string.IsNullOrWhiteSpace(rowKey))
                return string.Empty;

            return string.Join(" / ", rowKey
                .Split(global::Handlers.DataHandling.BoardDraftNaturalKeys.Separator)
                .Where(part => !string.IsNullOrWhiteSpace(part)));
        }
    }

    // ###########################################################################################
    // One scope setting that moved, named by the workbook's own column.
    // ###########################################################################################
    public sealed record ReviewScopeSetting(string Field, string Before, string After);

    // ###########################################################################################
    // Every scope setting that moved on one baseline.
    //
    // A record rather than a bare list so the thing being passed around says what it is - a caller
    // holding an IReadOnlyList<ReviewScopeSetting> has to remember which row it belongs to.
    // ###########################################################################################
    public readonly record struct ReviewScopeSettingsChange(IReadOnlyList<ReviewScopeSetting> Settings);
}
