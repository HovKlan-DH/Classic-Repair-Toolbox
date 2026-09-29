using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardViewContract and BoardViewRules - the part of board views (owner request, 2026-09-27) that
// CRT and CRT.Server both apply: which views still count, the time a view is stored at, how a
// sender's text is made fit to store, and what CRT does with each answer the server can give.
// ###########################################################################################
public sealed class BoardViewContractTests
{
    private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

    // A view counts from MaxAge ago up to MaxAhead into the future, both ends included.
    [Fact]
    public void A_view_counts_from_max_age_ago_up_to_max_ahead()
    {
        Assert.True(BoardViewRules.IsCountable(BoardViewContractTests.Now, BoardViewContractTests.Now));
        Assert.True(BoardViewRules.IsCountable(BoardViewContractTests.Now - BoardViewRules.MaxAge, BoardViewContractTests.Now));
        Assert.True(BoardViewRules.IsCountable(BoardViewContractTests.Now + BoardViewRules.MaxAhead, BoardViewContractTests.Now));

        Assert.False(BoardViewRules.IsCountable(BoardViewContractTests.Now - BoardViewRules.MaxAge - TimeSpan.FromSeconds(1), BoardViewContractTests.Now));
        Assert.False(BoardViewRules.IsCountable(BoardViewContractTests.Now + BoardViewRules.MaxAhead + TimeSpan.FromSeconds(1), BoardViewContractTests.Now));
    }

    // Stored in UTC, to the second, and never later than the server's own clock.
    [Fact]
    public void A_view_is_stored_in_utc_to_the_second_and_never_in_the_future()
    {
        var local = new DateTimeOffset(2026, 9, 27, 13, 30, 15, 750, TimeSpan.FromHours(2));

        Assert.Equal(new DateTimeOffset(2026, 9, 27, 11, 30, 15, TimeSpan.Zero), BoardViewRules.StoredTime(local, BoardViewContractTests.Now));
        Assert.Equal(TimeSpan.Zero, BoardViewRules.StoredTime(local, BoardViewContractTests.Now).Offset);
        Assert.Equal(BoardViewContractTests.Now, BoardViewRules.StoredTime(BoardViewContractTests.Now.AddHours(5), BoardViewContractTests.Now));
    }

    [Theory]
    [InlineData("Windows", 20, "Windows")]
    [InlineData("  Windows  ", 20, "Windows")]
    [InlineData("Win\r\ndows\t", 20, "Windows")]
    [InlineData("abcdefghij", 4, "abcd")]
    [InlineData("ab  cdef", 4, "ab")]
    [InlineData("", 20, null)]
    [InlineData("   ", 20, null)]
    [InlineData("\r\n", 20, null)]
    [InlineData(null, 20, null)]
    public void A_senders_text_loses_control_characters_and_is_cut_to_fit(string? value, int maxLength, string? expected)
    {
        Assert.Equal(expected, BoardViewRules.Clip(value, maxLength));
    }

    // ###########################################################################################
    // *** NEVER HALF A CHARACTER (code review, 2026-09-29). *** Clip cut at a UTF-16 index, so a
    // character outside the Basic Multilingual Plane (an emoji, some CJK) straddling the limit left a
    // lone high surrogate - which MariaDB's utf8mb4 refuses, failing the WHOLE batch's insert with a
    // 500. CRT then re-sends that identical batch every 15 minutes for ever, and every later view
    // from that machine queues behind it. A lone surrogate already in the text is dropped for the
    // same reason. U+1F600 is written as its two code units so the test file stays ASCII.
    // ###########################################################################################
    [Theory]
    [InlineData("abc\uD83D\uDE00", 4, "abc")]
    [InlineData("abc\uD83D\uDE00", 5, "abc\uD83D\uDE00")]
    public void A_senders_text_is_never_cut_or_left_in_the_middle_of_a_character(string value, int maxLength, string expected)
    {
        string? clipped = BoardViewRules.Clip(value, maxLength);

        Assert.Equal(expected, clipped);

        // Every character whole: a lone surrogate in a string makes its UTF-8 encoding fail.
        for (int index = 0; index < clipped!.Length; index++)
        {
            if (char.IsHighSurrogate(clipped[index]))
                Assert.True(index + 1 < clipped.Length && char.IsLowSurrogate(clipped[++index]));
            else
                Assert.False(char.IsLowSurrogate(clipped[index]));
        }
    }

    // A lone surrogate already in the text - half of a character, from a broken sender - is dropped
    // for the same reason. Built from character codes rather than as test data, because the test
    // runner may rewrite an unpaired surrogate when it serialises a data row.
    [Fact]
    public void A_lone_half_of_a_character_already_in_the_text_is_dropped()
    {
        string loneHigh = "a" + (char)0xD83D + "b";
        string loneLow = "a" + (char)0xDE00 + "b";
        string highAtTheEnd = "ab" + (char)0xD83D;

        Assert.Equal("ab", BoardViewRules.Clip(loneHigh, 20));
        Assert.Equal("ab", BoardViewRules.Clip(loneLow, 20));
        Assert.Equal("ab", BoardViewRules.Clip(highAtTheEnd, 20));
    }

    // ###########################################################################################
    // *** WHAT CRT DOES WITH EACH ANSWER. *** Accepted is DONE, and so is a refusal of THIS report
    // (400, 413, 415, 422) - it will not be accepted next time, and keeping it would block every
    // view behind it. Every other answer keeps it: none at all, a timeout, too many, a server fault,
    // and any other 4xx - 404 is a server without the route yet, 401/403/405 a proxy or firewall set
    // up wrong on the server's side. None of those says anything about the report, and dropping on
    // them would make every CRT throw its views away until the server was put right.
    // ###########################################################################################
    [Theory]
    [InlineData(200, BoardViewDelivery.Done)]
    [InlineData(202, BoardViewDelivery.Done)]
    [InlineData(204, BoardViewDelivery.Done)]
    [InlineData(400, BoardViewDelivery.Done)]
    [InlineData(413, BoardViewDelivery.Done)]
    [InlineData(415, BoardViewDelivery.Done)]
    [InlineData(422, BoardViewDelivery.Done)]
    [InlineData(401, BoardViewDelivery.TryLater)]
    [InlineData(403, BoardViewDelivery.TryLater)]
    [InlineData(405, BoardViewDelivery.TryLater)]
    [InlineData(410, BoardViewDelivery.TryLater)]
    [InlineData(404, BoardViewDelivery.TryLater)]
    [InlineData(408, BoardViewDelivery.TryLater)]
    [InlineData(429, BoardViewDelivery.TryLater)]
    [InlineData(500, BoardViewDelivery.TryLater)]
    [InlineData(502, BoardViewDelivery.TryLater)]
    [InlineData(503, BoardViewDelivery.TryLater)]
    [InlineData(0, BoardViewDelivery.TryLater)]
    [InlineData(301, BoardViewDelivery.TryLater)]
    public void Each_answer_either_finishes_a_report_or_keeps_it_for_later(int status, BoardViewDelivery expected)
    {
        Assert.Equal(expected, BoardViewContract.DeliveryFor(status));
    }

    // Ten seconds is the owner's number; a test names it so a change to it is a decision.
    [Fact]
    public void A_board_counts_after_ten_seconds_on_screen()
    {
        Assert.Equal(TimeSpan.FromSeconds(10), BoardViewRules.MinimumTimeOnScreen);
    }
}
