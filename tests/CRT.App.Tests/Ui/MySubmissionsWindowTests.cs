using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "my submissions" window (NewContributeStrategy.md Phase 4, task 6).
//
// The refresh is driven through StatusLookupForTests rather than a real SubmissionClient, so
// these tests never touch the network (test rule 6). What that seam replaces is the HTTP call
// alone - the decision about WHICH rows get asked about, the write-back and the rebuild are all
// the shipped code.
//
// SubmissionReceiptStore is static, so this class points it at a temp file via LoadFrom, the same
// seam SubmissionReceiptStoreTests uses. It cannot ALSO join the "SubmissionReceipts" collection,
// because a headless UI test has to be in "HeadlessUi" to share the dispatcher thread - a class
// may only be in one collection. Collections never run in parallel with each other
// (xunit.runner.json sets parallelizeTestCollections false), so the two cannot overlap anyway;
// that setting is what makes this safe, and it is the same reasoning WorkbooksListTests records
// for its own clash with "UserSettings".
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MySubmissionsWindowTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MySubmissionsWindowTests()
    {
        SubmissionReceiptStore.LoadFrom(Path.Combine(this.thisWorkspace.Root, "submissions.json"));
    }

    public void Dispose()
    {
        SubmissionReceiptStore.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static SubmissionReceipt Receipt(
        long id,
        string state = "",
        string boardId = "Commodore/C64/250407/Data C64 250407.xlsx")
        => new()
        {
            SubmissionId = id,
            UploadToken = "tok" + id,
            BoardId = boardId,
            Summary = "Fixed U8 pinout",
            SentUtc = new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero),
            LastKnownState = state
        };

    [Fact]
    public void With_nothing_sent_the_empty_state_shows_and_the_list_is_hidden()
    {
        UiTest.Run(() =>
        {
            var window = new MySubmissionsWindow();
            window.Initialize();

            Assert.Empty(window.Submissions);
            Assert.True(window.GetControl<TextBlock>("EmptyStateText").IsVisible);
            Assert.False(window.GetControl<ScrollViewer>("ListScrollViewer").IsVisible);

            // Nothing to check, so the button that would check it is off rather than offering an
            // action that can only be a no-op.
            Assert.False(window.GetControl<Button>("RefreshButton").IsEnabled);
        });
    }

    // ###########################################################################################
    // *** A DECISION THAT ARRIVED WITH NO COMMENT MUST BE DISMISSIBLE (owner report,
    // 2026-09-23). ***
    //
    // A maintainer approving good work usually writes nothing, so that row has no maintainer panel -
    // and both the "New" badge and the "Mark as read" button used to live INSIDE that panel. The
    // tab badge counted the row, correctly, and the window then offered no way whatever to clear
    // it: the project owner saw "2" and could find nothing to act on.
    //
    // The first attempt at this cleared such rows automatically when the window opened. That was
    // WRONG and the project owner said so: rows are tracked individually, and opening a list to check
    // on one submission must not silently mark another as seen. The markup had already recorded
    // that decision ("EXPLICIT DISMISS, not cleared by opening this window"). So the fix is a
    // dismiss control on the row itself, and opening the window still changes nothing.
    // ###########################################################################################
    [Fact]
    public void Opening_the_window_marks_NOTHING_as_read()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "published"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            // Still unread after simply looking at the list.
            Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());
            Assert.True(Assert.Single(window.Submissions).HasUnreadNews);
        });
    }

    [Fact]
    public void A_published_submission_with_no_comment_offers_its_own_dismiss()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "published"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            SubmissionListItem row = Assert.Single(window.Submissions);

            // There is news, there is no comment to carry the existing button, so the row's own
            // control is what has to be shown.
            Assert.True(row.HasUnreadNews);
            Assert.False(row.HasUnreadComment);
            Assert.True(row.HasUnreadDecisionOnly);
        });
    }

    [Fact]
    public void Dismissing_a_decision_only_row_clears_the_badge()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "published"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());

            Assert.Single(window.Submissions).MarkReadCommand.Execute(null);

            // The command posts through Dispatcher.UIThread.InvokeAsync, so the work has not run
            // yet - the same reason the forget test below pumps the dispatcher.
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
            Assert.False(Assert.Single(window.Submissions).HasUnreadNews);
        });
    }

    // Two rows, one dismissed: the other must stay unread. This is the "track them individually"
    // requirement stated as a test - a fix that cleared everything on open, or on any dismiss,
    // fails here.
    [Fact]
    public void Dismissing_ONE_row_leaves_the_others_unread()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "published"));
            SubmissionReceiptStore.Record(Receipt(43, "merged", "Commodore/C64/250425/Data.xlsx"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            Assert.Equal(2, SubmissionReceiptStore.UnreadCommentCount());

            window.Submissions.Single(row => row.SubmissionId == 42)
                .MarkReadCommand.Execute(null);

            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());
            Assert.False(window.Submissions.Single(row => row.SubmissionId == 42).HasUnreadNews);
            Assert.True(window.Submissions.Single(row => row.SubmissionId == 43).HasUnreadNews);
        });
    }

    // A row WITH an unread comment already carries a "New" badge and a button inside the maintainer
    // panel, so the row-level pair must stay hidden - otherwise the card says "New" twice about
    // the same thing.
    [Fact]
    public void A_row_with_an_unread_COMMENT_does_not_ALSO_show_the_row_level_dismiss()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "changes_requested"));
            SubmissionReceiptStore.UpdateState(
                42,
                "changes_requested",
                "Please add the region.",
                DateTimeOffset.UtcNow);

            var window = new MySubmissionsWindow();
            window.Initialize();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.True(row.HasUnreadComment);
            Assert.True(row.HasUnreadNews);
            Assert.False(row.HasUnreadDecisionOnly);
        });
    }

    // A submission still waiting for a maintainer is not news and must not be badged - otherwise
    // every contribution lights up the moment it is sent.
    [Fact]
    public void A_PENDING_submission_shows_no_dismiss_and_counts_nothing()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.False(row.HasUnreadNews);
            Assert.False(row.HasUnreadDecisionOnly);
            Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
        });
    }

    [Fact]
    public void A_recorded_submission_appears_as_a_row()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.Equal(42, row.SubmissionId);
            Assert.Equal("Submitted - awaiting feedback from a maintainer", row.StateText);
            Assert.Equal("Fixed U8 pinout", row.Summary);

            Assert.False(window.GetControl<TextBlock>("EmptyStateText").IsVisible);
            Assert.True(window.GetControl<ScrollViewer>("ListScrollViewer").IsVisible);
        });
    }

    // ###########################################################################################
    // The row names the BOARD, not the file. A board's identity is an ExcelDataFile key, and the
    // ".xlsx" on the end names a file the contributor never typed and, for a draft-only board,
    // one that does not even exist.
    // ###########################################################################################
    [Fact]
    public void A_row_names_the_board_rather_than_the_data_file()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42));

            var window = new MySubmissionsWindow();
            window.Initialize();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.Equal("Commodore C64 250407", row.BoardDisplayName);
            Assert.DoesNotContain(".xlsx", row.BoardDisplayName);
        });
    }

    // ###########################################################################################
    // *** A REAL RECEIPT CARRIES THE BOARD ID, NOT A WORKBOOK PATH (2026-09-27). *** SubmitDraftWindow
    // records identity.BoardId - "Commodore/C128/310378 Open128" - and the name used to drop the
    // last segment as though it were a file name, so the row read "Commodore C128" with the board
    // missing. The fixture above uses the old workbook-path shape no real receipt has, which is how
    // this went unseen; both shapes are named in full now.
    // ###########################################################################################
    [Fact]
    public void A_row_names_the_WHOLE_board_from_a_real_receipts_board_id()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(43, boardId: "Commodore/C128/310378 Open128"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            Assert.Equal("Commodore C128 310378 Open128", Assert.Single(window.Submissions).BoardDisplayName);
        });
    }

    // ###########################################################################################
    // THE REFRESH ONLY ASKS ABOUT ROWS THAT CAN STILL MOVE.
    //
    // A decided submission can never change state again, so asking about it is a request per row
    // per refresh that can only return what is already on screen. Asserted by counting the ids the
    // lookup was actually handed - a version that asked about everything still LOOKS right on
    // screen, which is exactly why this needs pinning rather than eyeballing.
    // ###########################################################################################
    [Fact]
    public void Refreshing_asks_only_about_submissions_that_are_not_yet_decided()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(1, "pending"));
            SubmissionReceiptStore.Record(Receipt(2, "rejected"));
            SubmissionReceiptStore.Record(Receipt(3, "abandoned"));
            SubmissionReceiptStore.Record(Receipt(4, string.Empty));

            var asked = new List<long>();

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) =>
            {
                asked.Add(id);
                return Task.FromResult<SubmissionStatus?>(new SubmissionStatus { Id = id, State = "pending" });
            };

            window.Initialize();
            window.RefreshForTests();

            // 1 is pending and 4 has never been checked; 2 and 3 are settled.
            Assert.Equal(new long[] { 1, 4 }, asked.OrderBy(id => id));
        });
    }

    // ###########################################################################################
    // The launch check stops asking about a submission in BETA a month after its decision (code
    // review, 2026-09-25). "Check for updates" does not - the contributor pressed it to ask, and
    // this is how they find out a board they sent long ago has since reached the source.
    // ###########################################################################################
    [Fact]
    public void Check_for_updates_still_asks_about_a_submission_long_in_BETA()
    {
        UiTest.Run(() =>
        {
            SubmissionReceipt old = Receipt(5, "merged");
            SubmissionReceiptStore.Record(new SubmissionReceipt
            {
                SubmissionId = old.SubmissionId,
                UploadToken = old.UploadToken,
                BoardId = old.BoardId,
                Summary = old.Summary,
                SentUtc = DateTimeOffset.UtcNow.AddYears(-1),
                DecidedUtc = DateTimeOffset.UtcNow.AddYears(-1),
                LastKnownState = "merged"
            });

            var asked = new List<long>();

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) =>
            {
                asked.Add(id);
                return Task.FromResult<SubmissionStatus?>(new SubmissionStatus { Id = id, State = "published" });
            };

            window.Initialize();
            window.RefreshForTests();

            Assert.Equal(new long[] { 5 }, asked);
        });
    }

    [Fact]
    public void A_refresh_writes_the_new_state_back_and_the_row_shows_it()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) => Task.FromResult<SubmissionStatus?>(
                new SubmissionStatus
                {
                    Id = id,
                    State = "rejected",
                    MaintainerComment = "The highlight for U8 is outside the image."
                });

            window.Initialize();
            window.RefreshForTests();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.Equal("Not accepted", row.StateText);
            Assert.Equal("The highlight for U8 is outside the image.", row.MaintainerComment);
            Assert.True(row.HasMaintainerComment);

            // And it survives a reload - the point of writing it back at all.
            Assert.Equal("rejected", SubmissionReceiptStore.All.Single().LastKnownState);
        });
    }

    // ###########################################################################################
    // AN UNREACHABLE SERVER LEAVES THE ROW ALONE rather than blanking it.
    //
    // GetStatusAsync answers null for anything it cannot reach. The row must then keep its last
    // known state: that cached value exists precisely so this window says something true with no
    // connection, and overwriting it with "unknown" on every failed refresh would destroy the only
    // information the user had.
    // ###########################################################################################
    [Fact]
    public void A_submission_the_server_cannot_answer_for_keeps_its_last_known_state()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (_, _, _) => Task.FromResult<SubmissionStatus?>(null);

            window.Initialize();
            window.RefreshForTests();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.Equal("Submitted - awaiting feedback from a maintainer", row.StateText);
            Assert.Equal("pending", SubmissionReceiptStore.All.Single().LastKnownState);

            // And the user is told, rather than left looking at a list that did not visibly change.
            TextBlock status = window.GetControl<TextBlock>("StatusText");
            Assert.True(status.IsVisible);
            Assert.Contains("could not be reached", status.Text);
        });
    }

    // ###########################################################################################
    // *** "NOT FOUND" IS AN ANSWER, NOT AN UNREACHABLE SERVER (code review, 2026-10-04). *** After a
    // reset of the contribution data, or a board deleted, every open receipt is answered 404. Those
    // were counted as "could not be reached", so the line said the server could not be reached
    // while the rows said "No longer on the server" - and the contributor went looking for a
    // network problem that was not there.
    // ###########################################################################################
    [Fact]
    public void Submissions_no_longer_on_the_server_are_counted_as_checked_not_as_unreachable()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(41, "pending"));
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) => throw new SubmissionNotFoundException(id);

            window.Initialize();
            window.RefreshForTests();

            TextBlock status = window.GetControl<TextBlock>("StatusText");
            Assert.Equal("Checked 2 contributions. 2 are no longer on the server.", status.Text);
            Assert.All(SubmissionReceiptStore.All, receipt => Assert.NotNull(receipt.NotFoundUtc));
        });
    }

    // One gone, one answered, one unreachable: each is said for what it is.
    [Fact]
    public void A_mix_of_gone_answered_and_unreachable_says_each_for_what_it_is()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(41, "pending"));
            SubmissionReceiptStore.Record(Receipt(42, "pending"));
            SubmissionReceiptStore.Record(Receipt(43, "pending"));

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) => id switch
            {
                41 => throw new SubmissionNotFoundException(id),
                42 => Task.FromResult<SubmissionStatus?>(new SubmissionStatus { Id = id, State = "pending" }),
                _ => Task.FromResult<SubmissionStatus?>(null)
            };

            window.Initialize();
            window.RefreshForTests();

            Assert.Equal(
                "Checked 2, but 1 could not be reached. Those rows show their last known state. 1 is no longer on the server.",
                window.GetControl<TextBlock>("StatusText").Text);
        });
    }

    // ###########################################################################################
    // *** "UPDATE CRT" IS SAID IN THE SERVER'S WORDS (code review, 2026-10-04). *** Not "the server
    // could not be reached", which sends the contributor looking for a network problem that is not
    // there. The rows keep their last known state, and nothing after the first answer is asked.
    // ###########################################################################################
    [Fact]
    public void An_update_CRT_answer_is_shown_in_the_servers_words_and_the_rows_keep_their_state()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(41, "pending"));
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var asked = new List<long>();
            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) =>
            {
                asked.Add(id);
                throw new ClientOutdatedException("Please update CRT to version [3.2.0] or newer.");
            };

            window.Initialize();
            window.RefreshForTests();

            Assert.Single(asked);
            Assert.All(window.Submissions, row => Assert.Equal("Submitted - awaiting feedback from a maintainer", row.StateText));

            TextBlock status = window.GetControl<TextBlock>("StatusText");
            Assert.True(status.IsVisible);
            Assert.Equal(SubmissionStatusRefresh.DescribeOutdated("Please update CRT to version [3.2.0] or newer."), status.Text);
        });
    }

    // Pressing a button and seeing nothing happen reads as a broken button. When every row is
    // already decided there is genuinely nothing to do, and saying so is the difference.
    [Fact]
    public void Refreshing_with_everything_already_decided_says_so_rather_than_appearing_to_do_nothing()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(1, "rejected"));
            SubmissionReceiptStore.Record(Receipt(2, "abandoned"));

            bool askedAnything = false;

            var window = new MySubmissionsWindow();
            window.StatusLookupForTests = (id, _, _) =>
            {
                askedAnything = true;
                return Task.FromResult<SubmissionStatus?>(new SubmissionStatus { Id = id, State = "pending" });
            };

            window.Initialize();
            window.RefreshForTests();

            Assert.False(askedAnything);

            TextBlock status = window.GetControl<TextBlock>("StatusText");
            Assert.True(status.IsVisible);
            Assert.Contains("already been decided", status.Text);
        });
    }

    // The maintainer-comment panel stays collapsed until there is a comment. Phase 5 is what fills
    // these in; until then every row would otherwise carry an empty labelled box.
    [Fact]
    public void A_row_with_no_maintainer_comment_does_not_offer_an_empty_comment_panel()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));

            var window = new MySubmissionsWindow();
            window.Initialize();

            Assert.False(Assert.Single(window.Submissions).HasMaintainerComment);
        });
    }

    // A never-checked receipt says so rather than claiming a state nobody reported.
    [Fact]
    public void A_submission_never_checked_shows_that_rather_than_a_made_up_state()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42));

            var window = new MySubmissionsWindow();
            window.Initialize();

            SubmissionListItem row = Assert.Single(window.Submissions);

            Assert.Equal("Not checked yet", row.StateText);
            Assert.False(row.HasCheckedText);
        });
    }

    // ------------------------------------------------------------------ The card's own markup

    // ###########################################################################################
    // THE CARD IS BUILT FROM A DataTemplate, so these are the only tests that can see what it
    // actually renders - the row model exposes CheckedText whether or not the card binds it.
    //
    // Shown through a real layout pass (UiTest's window + measure), because an ItemsControl that
    // has never been laid out has generated no containers at all and every query returns nothing,
    // which would make each of these pass vacuously.
    // ###########################################################################################
    private static IReadOnlyList<TextBlock> CardTextBlocks(MySubmissionsWindow window)
    {
        var items = window.GetControl<ItemsControl>("SubmissionsItemsControl");

        window.Measure(new Avalonia.Size(900, 900));
        window.Arrange(new Avalonia.Rect(0, 0, 900, 900));
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        return items.GetVisualDescendants().OfType<TextBlock>().ToList();
    }

    // ###########################################################################################
    // *** "Last checked" IS GONE FROM THE CARD (owner request, 2026-09-23). ***
    //
    // It was this computer's bookkeeping - when CRT last asked the server - and said nothing about
    // the submission. Asserted against the RENDERED card rather than against the row model, since
    // SubmissionListItem still exposes CheckedText and a model-level check would pass while the
    // line was still on screen.
    // ###########################################################################################
    [Fact]
    public void The_card_does_NOT_show_a_last_checked_date()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));
            SubmissionReceiptStore.UpdateState(
                42, "changes_requested", "Please change", DateTimeOffset.UtcNow,
                new DateTimeOffset(2026, 9, 22, 9, 0, 0, TimeSpan.Zero));

            var window = new MySubmissionsWindow();
            window.Initialize();

            IReadOnlyList<TextBlock> blocks = MySubmissionsWindowTests.CardTextBlocks(window);

            // The card must have rendered at all, or the assertion below means nothing.
            Assert.Contains(blocks, block => block.Text?.StartsWith("Sent ", StringComparison.Ordinal) == true);

            Assert.DoesNotContain(
                blocks,
                block => block.Text?.Contains("Last checked", StringComparison.Ordinal) == true);
        });
    }

    // ###########################################################################################
    // The maintainer panel's heading, verbatim (owner request, 2026-09-23): "Feedback from
    // maintainer", with the date on its OWN line beneath it rather than beside it.
    //
    // The two-line shape is what was asked for - side by side the pair read as a caption strip
    // rather than as the heading of something to be read - so the DATE BEING ITS OWN TextBlock is
    // the thing worth pinning, not merely that both strings appear somewhere.
    // ###########################################################################################
    [Fact]
    public void The_maintainer_panel_is_headed_by_the_label_with_the_date_on_its_own_line()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(42, "pending"));
            SubmissionReceiptStore.UpdateState(
                42, "changes_requested", "Please change", DateTimeOffset.UtcNow,
                new DateTimeOffset(2026, 9, 22, 9, 0, 0, TimeSpan.Zero));

            var window = new MySubmissionsWindow();
            window.Initialize();

            IReadOnlyList<TextBlock> blocks = MySubmissionsWindowTests.CardTextBlocks(window);

            Assert.Contains(blocks, block => block.Text == "Feedback from maintainer");

            // The old wording, which must not survive anywhere on the card.
            Assert.DoesNotContain(blocks, block => block.Text == "From the maintainer");

            // Its OWN block, carrying the date alone - not a label with the date appended.
            Assert.Contains(
                blocks,
                block => block.Text?.StartsWith("Replied ", StringComparison.Ordinal) == true
                    && !block.Text.Contains("maintainer", StringComparison.Ordinal));
        });
    }

    // ###########################################################################################
    // *** "CHECK FOR UPDATES" WAITS UNDER THE OVERLAY, AND GIVES UP AFTER THE LIMIT (2026-09-28). ***
    // The overlay is up while the server is asked; a server that never answers is let go when the
    // two minutes pass (the test's clock), the overlay lifts, and the line says the rows show what
    // was checked before that - nothing lost, and nothing claimed.
    // ###########################################################################################
    [Fact]
    public async Task A_check_with_no_answer_lifts_the_overlay_at_the_limit_and_says_what_is_shown()
    {
        await UiTest.RunAsync(async () =>
        {
            SubmissionReceiptStore.Record(Receipt(7, "pending"));

            var window = new MySubmissionsWindow();
            BusyOverlay overlay = BusyOverlay.For(window)!;

            var limit = new TaskCompletionSource();
            overlay.LimitOverrideForTests = _ => limit.Task;

            bool busyWhileAsking = false;

            window.StatusLookupForTests = (_, _, token) =>
            {
                busyWhileAsking = overlay.IsBusy;

                // The two minutes pass while this request is still out - and it never answers.
                limit.TrySetResult();
                return Task.Delay(Timeout.Infinite, token).ContinueWith<SubmissionStatus?>(_ => null, TaskScheduler.Default);
            };

            window.Initialize();

            await (Task)typeof(MySubmissionsWindow)
                .GetMethod("RefreshAsync", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .Invoke(window, null)!;

            Assert.True(busyWhileAsking);
            Assert.False(overlay.IsBusy);
            Assert.Equal(CrtWaitWording.SubmissionsNoAnswer, window.FindControl<TextBlock>("StatusText")!.Text);
        });
    }
}
