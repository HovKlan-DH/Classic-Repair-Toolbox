using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Tests for OneSubmissionInBeta - the two sentences said when a board already waits in BETA for
// the stable source: refusing a second approval (BusyMessage) and refusing a change on the Boards
// screen (NoChangeMessage). The server sends them and the Maintainer tab shows them as they are.
// ###########################################################################################
public sealed class OneSubmissionInBetaTests
{
    private const string C64 = "Commodore/C64/250407";

    // ###########################################################################################
    // *** NEITHER SENTENCE TELLS ITS READER TO PUBLISH TO STABLE (code review, 2026-10-05). *** While
    // only the administrator publishes to the stable source, a maintainer refused with "publish that
    // one to stable" was offered the one thing their screen greys out. The sentences say what has to
    // happen first, which is true for whoever reads them.
    // ###########################################################################################
    [Fact]
    public void Neither_refusal_tells_a_maintainer_to_publish_to_stable_themselves()
    {
        Assert.Equal(
            $"{C64} already has a submission in BETA that has not gone to the stable source. A board takes one " +
            "submission into BETA at a time: that one has to be published to stable or pushed back to the queue " +
            $"(under {MaintainerScreenWording.BetaQueueQuoted}) first.",
            OneSubmissionInBeta.BusyMessage(C64));

        Assert.Equal(
            $"{C64} is waiting in BETA for the stable source, so no change can be made to it here. It has to be " +
            $"published to stable, pushed back or rejected under {MaintainerScreenWording.BetaQueueQuoted} first.",
            OneSubmissionInBeta.NoChangeMessage(C64));
    }

    // A blank board id still makes a sentence, and the id is trimmed when there is one.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_missing_board_id_reads_as_this_board(string? boardId)
    {
        Assert.StartsWith("This board already has", OneSubmissionInBeta.BusyMessage(boardId), StringComparison.Ordinal);
        Assert.StartsWith("This board is waiting", OneSubmissionInBeta.NoChangeMessage(boardId), StringComparison.Ordinal);
    }

    [Fact]
    public void The_board_id_is_trimmed()
    {
        Assert.StartsWith($"{C64} already has", OneSubmissionInBeta.BusyMessage($"  {C64} "), StringComparison.Ordinal);
    }
}
