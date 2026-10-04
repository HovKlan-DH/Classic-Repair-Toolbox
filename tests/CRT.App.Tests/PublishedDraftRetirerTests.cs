using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PublishedDraftRetirer - the sequencing around deleting a draft whose work has been published
// (code review, 2026-09-25). WHETHER a draft may go is DraftRetirement's rule, pinned in
// CRT.Data.Tests; these pin HOW the app goes about it:
//
//   - the slow check runs off the calling (UI) thread - it used to freeze the window;
//   - a draft edited between the check and the delete is kept;
//   - a draft with unsaved table edits is kept, since those edits are not on disk;
//   - the caller's cache-clearing runs for every folder a delete was attempted on, finished or
//     not - a board left cached from a deleted folder kept showing drafted rows.
//
// Everything with a side effect is a delegate, so no test here deletes through the app's real
// DraftManager or touches the user's folders.
// ###########################################################################################
public sealed class PublishedDraftRetirerTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private const string SystemKey = "Commodore/C64/250407/Data.xlsx";

    // A draft folder holding one file, and a candidate stamped as it is now.
    private RetirableDraft CandidateOnDisk()
    {
        string folder = Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C64", "250407");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "Data.xlsx"), "workbook");

        return new RetirableDraft(PublishedDraftRetirerTests.SystemKey, folder, DraftRetirement.FolderStamp(folder));
    }

    // ###########################################################################################
    // *** CALLED FROM A THREAD OF ITS OWN THAT STAYS BUSY, as the UI thread does. *** Called from
    // the test's own thread - a POOL thread under xunit v3 - it failed now and then (2026-10-03, a
    // full run): the test's await handed its thread back to the pool, which could then run the
    // check on that very thread, and the two ids matched although nothing was wrong.
    // ###########################################################################################
    [Fact]
    public void The_check_runs_OFF_the_calling_thread()
    {
        // It parses two workbooks and reads every file per draft. Run on the UI thread, as it
        // used to be inside Dispatcher.UIThread.Post, that froze the window.
        int callingThread = 0;
        int? checkingThread = null;

        var caller = new Thread(() =>
        {
            callingThread = Environment.CurrentManagedThreadId;

            PublishedDraftRetirer.FindAsync(
                [new SubmissionReceipt { SubmissionId = 1, SystemId = PublishedDraftRetirerTests.SystemKey, LastKnownState = "published" }],
                _ =>
                {
                    checkingThread = Environment.CurrentManagedThreadId;
                    return null;
                }).GetAwaiter().GetResult();
        });

        caller.Start();
        Assert.True(caller.Join(TimeSpan.FromSeconds(30)), "the check never finished");

        Assert.NotNull(checkingThread);
        Assert.NotEqual(callingThread, checkingThread);
    }

    [Fact]
    public void A_draft_found_retirable_and_unchanged_is_deleted_and_its_cache_cleared()
    {
        RetirableDraft candidate = this.CandidateOnDisk();
        var discarded = new List<string>();
        var cleared = new List<string>();

        DraftRetirementOutcome outcome = PublishedDraftRetirer.Retire(
            [candidate],
            isInUse: _ => false,
            discard: systemId => { discarded.Add(systemId); return true; },
            afterDiscard: cleared.Add);

        Assert.Equal([PublishedDraftRetirerTests.SystemKey], discarded);
        Assert.Equal([PublishedDraftRetirerTests.SystemKey], cleared);
        Assert.Equal([PublishedDraftRetirerTests.SystemKey], outcome.Retired);
        Assert.Equal([PublishedDraftRetirerTests.SystemKey], outcome.Touched);
    }

    // ###########################################################################################
    // *** THE GAP BETWEEN CHECK AND DELETE. *** The check now runs on a pool thread and takes a
    // while, so a save can land in the draft before the delete. The folder's stamp from before the
    // check must still hold, or the draft is kept.
    // ###########################################################################################
    [Fact]
    public void A_draft_that_CHANGED_after_the_check_is_kept()
    {
        RetirableDraft candidate = this.CandidateOnDisk();

        File.WriteAllText(Path.Combine(candidate.Folder, "saved-meanwhile.txt"), "new work");

        bool discardCalled = false;

        DraftRetirementOutcome outcome = PublishedDraftRetirer.Retire(
            [candidate],
            isInUse: _ => false,
            discard: _ => discardCalled = true,
            afterDiscard: _ => { });

        Assert.False(discardCalled);
        Assert.Empty(outcome.Touched);
    }

    [Fact]
    public void A_draft_with_UNSAVED_table_edits_is_kept()
    {
        // Those edits live in memory only, so a folder identical to the published board proves
        // nothing about them.
        RetirableDraft candidate = this.CandidateOnDisk();

        bool discardCalled = false;

        DraftRetirementOutcome outcome = PublishedDraftRetirer.Retire(
            [candidate],
            isInUse: _ => true,
            discard: _ => discardCalled = true,
            afterDiscard: _ => { });

        Assert.False(discardCalled);
        Assert.Empty(outcome.Touched);
    }

    [Fact]
    public void A_delete_that_FAILS_still_clears_the_cache_and_is_reported()
    {
        // A half-finished delete (a file locked in Excel) still removed files the cached board
        // describes, so the cache must go either way - and the board on screen be reloaded, which
        // is what Touched tells Main.
        RetirableDraft candidate = this.CandidateOnDisk();
        var cleared = new List<string>();

        DraftRetirementOutcome outcome = PublishedDraftRetirer.Retire(
            [candidate],
            isInUse: _ => false,
            discard: _ => false,
            afterDiscard: cleared.Add);

        Assert.Equal([PublishedDraftRetirerTests.SystemKey], cleared);
        Assert.Empty(outcome.Retired);
        Assert.Equal([PublishedDraftRetirerTests.SystemKey], outcome.Failed);
        Assert.Equal([PublishedDraftRetirerTests.SystemKey], outcome.Touched);
    }
}
