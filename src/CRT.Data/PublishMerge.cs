using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Turns a submission into the board that will actually be written
    // (NewContributeStrategy.md Phase 5, task 6 - "the caller that assembles the merged
    // BoardData").
    //
    // *** "MERGE" IS THE WRONG WORD FOR WHAT THIS DOES, and believing otherwise corrupts boards.
    // *** The manifest is BY CONTRACT the complete intended state - SubmissionManifest.Files says
    // it outright: "a file present in the base revision but ABSENT here is a deletion", and the
    // rows follow the same rule. So the submitted rows REPLACE the published ones wholesale.
    //
    // Anything that tried to combine the two - keeping a published row the submission omitted -
    // would resurrect every row a contributor deliberately deleted, on every publish, and nothing
    // would ever report it. The contributor would delete a wrong component, watch it get approved,
    // and find it still there.
    //
    // What this class genuinely decides is the small set of things the rows do NOT settle:
    //
    //   - the REVISION DATE, which an older client may not send at all;
    //   - the PUBLISHED REVISION, which is a contract with the client rather than a free choice;
    //   - the CALIBRATIONS, which are not part of BoardData and travel separately.
    //
    // Pure, and in CRT.Data rather than the server, because the same "what would this submission
    // make the board look like" question is what the review screen asks - see
    // SubmissionBoardData.FromManifest, which this deliberately does not duplicate but extends.
    // ###########################################################################################
    public static class PublishMerge
    {
        // ###########################################################################################
        // The board to write.
        //
        // published may be null - that is a NEW SYSTEM, the ordinary case for the highest-risk kind
        // of submission, not an error.
        //
        // *** THE REVISION DATE IS THE ONE FIELD THAT FALLS BACK. *** Every other section is taken
        // from the submission alone. The date falls back to the published one because an older
        // client does not send it at all (the field was added 2026-09-22), and blanking it would
        // be far worse than leaving it: it is what a contributor's draft records as its
        // BaseRevision, so an empty one makes every existing draft look as though it were based on
        // nothing and the drift check has no answer to give.
        // ###########################################################################################
        public static BoardData Build(SubmissionManifest manifest, BoardData? published)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            SubmissionRows rows = manifest.Rows ?? new SubmissionRows();

            string revisionDate = !string.IsNullOrWhiteSpace(rows.RevisionDate)
                ? rows.RevisionDate.Trim()
                : published?.RevisionDate ?? string.Empty;

            // Every list is COPIED, never aliased. The manifest is reloaded per request and the
            // board is handed to a writer; sharing instances means a later mutation of one
            // silently changes the other.
            return new BoardData
            {
                RevisionDate = revisionDate,
                Schematics = [.. rows.Schematics],
                Components = [.. rows.Components],
                ComponentImages = [.. rows.ComponentImages],
                ComponentHighlights = [.. rows.ComponentHighlights],
                ComponentLocalFiles = [.. rows.ComponentLocalFiles],
                ComponentLinks = [.. rows.ComponentLinks],
                BoardLocalFiles = [.. rows.BoardLocalFiles],
                BoardLinks = [.. rows.BoardLinks],
                Credits = [.. rows.Credits],
                KiCadImportantSignals = [.. rows.KiCadImportantSignals]
            };
        }

        // ###########################################################################################
        // What `systems.current_revision` becomes.
        //
        // *** IT IS THE BOARD'S OWN REVISION DATE, and that is a CONTRACT WITH THE CLIENT rather
        // than a choice available here. *** DraftBaseRevision.EnsureBaseRevision stamps a draft's
        // BaseRevision from the official board's REVISION DATE, and the submission sends that
        // value back as its BaseRevision to be diffed against. If the server recorded anything
        // else - a counter, a timestamp, a hash - every contributor would be re-basing against a
        // value their drafts never carry, the drift warning would fire on every board forever, and
        // the "what changed officially since you started" screen would have nothing to compare.
        //
        // A method rather than a bare property read, so the reasoning above has somewhere to live
        // and the next person to want a counter here finds out why not.
        // ###########################################################################################
        public static string RevisionOf(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.RevisionDate?.Trim() ?? string.Empty;
        }

        // ###########################################################################################
        // The KiCad calibrations to write into the sidecar.
        //
        // *** NOT PART OF BoardData, WHICH IS WHY THIS EXISTS. *** Calibrations live only in the
        // `.json` sidecar, under their own root; BoardData has no section for them. A caller that
        // looked for them on the merged board would find nothing and publish a board with the
        // contributor's calibration work silently removed - the same class of loss the sidecar
        // writer was added to prevent.
        //
        // Returns an EMPTY list rather than null, because PublishExecutor refuses null outright
        // (a null there almost always means a caller failed to read something rather than a board
        // genuinely having none).
        //
        // The CLIENT fills this via KiCadCalibrationDraftWriter.CollectForSubmission, passed to
        // SubmissionManifestBuilder alongside the board. That was wired up on 2026-09-22; before
        // it, the list arrived empty from every real submission and a contributor's calibration
        // work silently never travelled.
        // ###########################################################################################
        public static IReadOnlyList<KiCadCalibrationEntry> CalibrationsOf(SubmissionManifest manifest)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return [.. (manifest.Rows ?? new SubmissionRows()).KiCadCalibrations];
        }
    }
}
