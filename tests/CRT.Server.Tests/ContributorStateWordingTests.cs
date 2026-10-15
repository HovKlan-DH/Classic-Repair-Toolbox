using System.Reflection;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests;

// ###########################################################################################
// EVERY STATE THE SERVER CAN TELL A CONTRIBUTOR, AND CRT'S WORDS FOR IT (code review, 2026-10-13).
//
// The states are plain strings on the wire, so the API surface freeze cannot see a new one - and
// the one cross-check there was covered a single value. A state the server adds and CRT does not
// know reads "Reported as [on_hold]" in "My submissions" and is asked about forever. This test
// holds the two halves together: every state the server can report (its database states, and the
// two it derives for contributors) has CRT words, and is classified as CRT means it.
// ###########################################################################################
public sealed class ContributorStateWordingTests
{
    // Every public const string on SubmissionState, plus the two contributor-facing words.
    private static IEnumerable<string> ServerStates() =>
        typeof(SubmissionState)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(field => field.IsLiteral && field.FieldType == typeof(string))
            .Select(field => (string)field.GetRawConstantValue()!)
            .Concat([ProductionPromotionRules.PublishedState, ProductionPromotionRules.ReturnedState]);

    public static TheoryData<string> States => new(ServerStates());

    [Theory]
    [MemberData(nameof(States))]
    public void Every_state_the_server_reports_has_words_in_CRT(string state)
    {
        string words = SubmissionReceiptPresenter.DescribeState(state);

        Assert.DoesNotContain("Reported as", words);
        Assert.False(string.IsNullOrWhiteSpace(words));
    }

    [Theory]
    [MemberData(nameof(States))]
    public void Every_state_the_server_reports_is_known_to_CRT_as_open_or_final(string state)
    {
        // Open states are listed; final ones answer "not open". Anything else is a word CRT would
        // only put on the slow re-check schedule.
        bool open = SubmissionReceiptPresenter.IsStillOpen(state);

        Assert.True(!open || SubmissionReceiptPresenter.KnownOpenStates.Contains(state), $"[{state}] is open but not known");
    }

    [Fact]
    public void The_list_is_not_empty()
    {
        Assert.Contains("pending", ServerStates());
        Assert.Contains("published", ServerStates());
        Assert.Contains("returned", ServerStates());
    }
}
