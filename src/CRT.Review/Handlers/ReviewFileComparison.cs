using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace CRT.Review.Handlers
{
    // ###########################################################################################
    // EVERY FILE a submission would change on the server, not only the ones that can be drawn
    // (security review, 2026-09-25).
    //
    // *** THE REVIEW SCREEN USED TO LIST IMAGES AND NOTHING ELSE. *** A PDF, a text file, a file
    // no row used or a file belonging to another board was shown nowhere, so an administrator
    // approved it without once seeing its name. ReviewImageComparison still draws the pictures;
    // this is the complete written list above them, and it says for each file what a reviewer
    // needs to weigh:
    //
    //   - whether it is NEW on the server or REPLACES something already published there;
    //   - whether it lives in a SHARED folder, where a change reaches every board that cites it;
    //   - whether it belongs to ANOTHER BOARD entirely, or is used by NO ROW at all.
    //
    // The server refuses the last two outright now (SubmissionFileRules), so they should never
    // appear. They are still flagged: this screen must not depend on the server having been right.
    //
    // A file the old board cited and the submission does not is listed as NO LONGER USED and stays
    // on the server - unless nothing else uses it either, when publishing REMOVES it (2026-09-25).
    // Which of the two is the server's FileRemovalPreview, the list the approval sends back.
    //
    // Pure, so what the reviewer is told is unit tested rather than checked by eye.
    // ###########################################################################################
    public static class ReviewFileComparison
    {
        // ###########################################################################################
        // The lines to show, "no longer used" first (the least visible change), then additions and
        // replacements in path order. A file byte-identical to what is already published is not a
        // change and is left out - a rows-only edit re-lists every file the board has.
        // ###########################################################################################
        public static IReadOnlyList<ReviewFileLine> Plan(
            IReadOnlyList<SubmittedFileFact>? submitted,
            IReadOnlyCollection<string>? publishedFiles,
            FileRemovalPreview? removals = null)
        {
            var lines = new List<ReviewFileLine>();

            // With no facts there is nothing to say "no longer used" against - an older server, or
            // a payload that could not be loaded. Saying every published file had gone would be a
            // loud, false alarm.
            if (submitted is null || submitted.Count == 0)
                return lines;

            var submittedPaths = new HashSet<string>(submitted.Select(fact => fact.Path), StringComparer.Ordinal);

            foreach (string path in (publishedFiles ?? [])
                .Where(path => !submittedPaths.Contains(path))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(path => path, StringComparer.Ordinal))
            {
                ReviewFileChange change = FileRemovalWording.IsRemoved(removals, path)
                    ? ReviewFileChange.Removed
                    : ReviewFileChange.NoLongerUsed;

                lines.Add(new ReviewFileLine(path, change, Scope: null, IsReferenced: false));
            }

            // A removed file the published board did not name in that exact spelling still gets
            // its own line: the list shown must be the whole list the server removes.
            foreach (string path in removals?.Files ?? [])
            {
                if (!lines.Any(line => string.Equals(line.Path, path, StringComparison.OrdinalIgnoreCase)))
                    lines.Add(new ReviewFileLine(path, ReviewFileChange.Removed, Scope: null, IsReferenced: false));
            }

            foreach (SubmittedFileFact fact in submitted.OrderBy(fact => fact.Path, StringComparer.Ordinal))
            {
                if (fact.IsUnchanged)
                    continue;

                lines.Add(new ReviewFileLine(
                    fact.Path,
                    fact.ExistsOnServer ? ReviewFileChange.Replaced : ReviewFileChange.Added,
                    fact.Scope,
                    fact.IsReferenced));
            }

            return lines;
        }

        // ###########################################################################################
        // "Replaced: Commodore/Shared files/7805.png" - the headline of one line.
        // ###########################################################################################
        public static string Describe(ReviewFileLine line)
        {
            ArgumentNullException.ThrowIfNull(line);

            string what = line.Change switch
            {
                ReviewFileChange.Added => "Added",
                ReviewFileChange.Replaced => "Replaced",
                ReviewFileChange.NoLongerUsed => "No longer used (stays on the server)",
                ReviewFileChange.Removed => "REMOVED from the server (nothing uses it any more)",
                _ => "Changed"
            };

            return $"{what}: {line.Path}";
        }

        // ###########################################################################################
        // What the reviewer should weigh about this file, one sentence each, most serious first.
        // Empty for an ordinary file in the board's own folder that a row uses.
        // ###########################################################################################
        public static IReadOnlyList<string> Warnings(ReviewFileLine line)
        {
            ArgumentNullException.ThrowIfNull(line);

            var warnings = new List<string>();

            if (line.Change is ReviewFileChange.NoLongerUsed or ReviewFileChange.Removed)
                return warnings;

            if (line.Scope == SubmissionFileScope.Foreign)
                warnings.Add("BELONGS TO ANOTHER BOARD - a submission should never change this.");

            if (!line.IsReferenced)
                warnings.Add("NO ROW USES THIS FILE - nothing in the board would show it.");

            if (line.Scope == SubmissionFileScope.ManufacturerShared)
                warnings.Add("Shared folder: every board of this manufacturer that uses it will change.");

            if (line.Scope == SubmissionFileScope.GenericShared)
                warnings.Add("Shared folder: every board that uses it, of any manufacturer, will change.");

            return warnings;
        }
    }

    // ###########################################################################################
    // One file on the list. Scope is null for a "no longer used" line, which carries no submitted
    // bytes to classify.
    // ###########################################################################################
    public sealed record ReviewFileLine(
        string Path,
        ReviewFileChange Change,
        SubmissionFileScope? Scope,
        bool IsReferenced);

    public enum ReviewFileChange
    {
        Added,
        Replaced,
        NoLongerUsed,

        // No longer used by this board AND by anything else: publishing removes it (2026-09-25).
        Removed
    }
}
