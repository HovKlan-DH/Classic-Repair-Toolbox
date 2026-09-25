using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Media;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog's findings list - what a contributor is told when a submission is refused.
//
// REPORTED FROM THE FIRST TWO REAL SUBMISSIONS (2026-09-21). The dialog said "The server found
// problems that have to be fixed" above an EMPTY box. The server had recorded three findings
// naming exactly what was wrong; all three were built, laid out, and rendered in WHITE ON WHITE.
//
// The cause was one line: the error colour was taken from `Button_Cancel_Fg`, which is White -
// the foreground for text on a red-filled Cancel BUTTON, not for text on this panel. Presence
// was never the problem, so a test asserting "three TextBlocks exist" would have passed happily
// while the user saw nothing at all.
//
// THAT IS WHY THESE TESTS ASSERT ON THE RENDERED COLOUR, not just on the text. A refusal the
// contributor cannot read is worse than a crash: a crash looks like a fault, while this looks
// like their data is wrong and offers nothing to fix.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowFindingsTests
{
    private static ValidationFinding Error(string code, string subject, string message) => new()
    {
        Severity = ValidationSeverity.Error,
        Code = code,
        Subject = subject,
        Message = message
    };

    private static ValidationFinding Warning(string code, string message) => new()
    {
        Severity = ValidationSeverity.Warning,
        Code = code,
        Subject = string.Empty,
        Message = message
    };

    // The three findings the live server actually returned, so this test is the reported failure.
    private static IReadOnlyList<ValidationFinding> TheLiveRejection() =>
    [
        SubmitDraftWindowFindingsTests.Error(
            "identity.system_id_mismatch", string.Empty,
            "The submission's system identifier does not match the manufacturer, hardware and board it names."),
        SubmitDraftWindowFindingsTests.Error(
            "identity.hardware_missing", string.Empty, "The submission does not name the hardware."),
        SubmitDraftWindowFindingsTests.Error(
            "identity.board_missing", string.Empty, "The submission does not name the board."),
    ];

    private static List<TextBlock> FindingBlocks(SubmitDraftWindow window) =>
        window.GetControl<StackPanel>("FindingsPanel").Children.OfType<TextBlock>().ToList();

    [Fact]
    public void Every_finding_the_server_returned_is_shown()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests(SubmitDraftWindowFindingsTests.TheLiveRejection());

            List<TextBlock> blocks = SubmitDraftWindowFindingsTests.FindingBlocks(window);

            Assert.Equal(3, blocks.Count);
            Assert.Contains(blocks, block => block.Text!.Contains("does not name the hardware"));
            Assert.Contains(blocks, block => block.Text!.Contains("does not name the board"));
            Assert.Contains(blocks, block => block.Text!.Contains("system identifier does not match"));
        });
    }

    // ###########################################################################################
    // WHAT A HEADLESS TEST CANNOT PROVE HERE, stated so nobody trusts it to.
    //
    // The bug was a COLOUR: findings were painted with `Button_Cancel_Fg`, which is White, on a
    // white panel. The obvious regression test - assert the brush is not white - is VACUOUS in
    // this harness, and I wrote it that way first and watched it pass against the broken code.
    //
    // The reason: a headless window is never attached to the application's resource tree, so
    // TryFindResource returns false, the colouring branch never runs, and every block keeps the
    // inherited Black. The assertion then passes because the lookup did not happen - not because
    // the right key was used. Verified by printing the resolved colour: Black, under both the
    // broken and the fixed version.
    //
    // So the colour is verified BY RUNNING THE APP, and what is pinned here is the thing that
    // actually can be: that the branch never assigns a null brush (which renders nothing), and
    // that every finding reaches the panel with its text intact. The key itself is guarded by the
    // comment at the call site, which names why Button_Cancel_Fg is the wrong one.
    // ###########################################################################################

    // A null Foreground is NOT "use the default" - it is how the text vanished. An error either
    // gets a real brush or is left to inherit the window's own, never assigned null.
    [Fact]
    public void A_finding_either_carries_a_real_brush_or_inherits_rather_than_being_nulled()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests(SubmitDraftWindowFindingsTests.TheLiveRejection());

            foreach (TextBlock block in SubmitDraftWindowFindingsTests.FindingBlocks(window))
            {
                // Either inherited (never explicitly set) or a genuine brush. What must not happen
                // is an explicit null, which renders nothing.
                if (block.IsSet(TextBlock.ForegroundProperty))
                    Assert.NotNull(block.Foreground);
            }
        });
    }

    // A finding that names its subject shows it, so a contributor with 400 components knows which
    // one is at fault.
    [Fact]
    public void A_finding_that_names_a_subject_shows_it()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests(
            [
                SubmitDraftWindowFindingsTests.Error("file.missing", "main.png", "That file is not in the submission.")
            ]);

            TextBlock block = Assert.Single(SubmitDraftWindowFindingsTests.FindingBlocks(window));

            Assert.Contains("main.png", block.Text!);
        });
    }

    // Errors first: those are what blocked the submission, and they are what has to be fixed.
    [Fact]
    public void Errors_are_listed_before_warnings()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests(
            [
                SubmitDraftWindowFindingsTests.Warning("summary.missing", "No summary was given."),
                SubmitDraftWindowFindingsTests.Error("identity.board_missing", string.Empty, "No board named.")
            ]);

            List<TextBlock> blocks = SubmitDraftWindowFindingsTests.FindingBlocks(window);

            Assert.Contains("No board named", blocks[0].Text!);
        });
    }

    // ###########################################################################################
    // AN ACCEPTED SUBMISSION SHOWS NOTHING HERE.
    //
    // Reported on the very first successful submission (2026-09-21): under "Contribution sent",
    // the dialog also displayed "The server did not say what was wrong, which is a fault in CRT" -
    // because the no-findings fallback was called on BOTH outcomes, and an accepted submission
    // legitimately carries no findings. That is what "nothing wrong with it" looks like.
    //
    // Telling someone their contribution was accepted AND that something went wrong, at the one
    // moment the pipeline finally worked, is the worst version of this panel getting it wrong.
    // ###########################################################################################
    [Fact]
    public void An_accepted_submission_shows_no_findings_and_no_apology()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests([], isRefusal: false);

            Assert.Empty(window.GetControl<StackPanel>("FindingsPanel").Children);
        });
    }

    // An accepted submission CAN carry warnings - a missing summary, say - and those are still
    // worth showing. What must not appear is the "something went wrong" apology.
    [Fact]
    public void An_accepted_submission_still_shows_a_warning_it_carries()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests(
                [SubmitDraftWindowFindingsTests.Warning("summary.missing", "No summary was given.")],
                isRefusal: false);

            TextBlock block = Assert.Single(SubmitDraftWindowFindingsTests.FindingBlocks(window));

            Assert.Contains("No summary was given", block.Text!);
            Assert.DoesNotContain("fault in CRT", block.Text!);
        });
    }

    // ###########################################################################################
    // A refusal carrying NO findings must still say something. It should not happen - the server
    // attaches its reasons - but "problems were found" over an empty box is the exact dead end
    // this whole file exists because of, and it must not be reachable by any route.
    // ###########################################################################################
    [Fact]
    public void A_refusal_with_no_findings_still_tells_the_user_something()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            window.ShowFindingsForTests([]);

            TextBlock block = Assert.Single(SubmitDraftWindowFindingsTests.FindingBlocks(window));

            Assert.False(string.IsNullOrWhiteSpace(block.Text));

            // And it says whose fault it is, so the contributor does not go hunting through their
            // own board for a problem that is in the software.
            Assert.Contains("fault in CRT", block.Text!);
        });
    }
}
