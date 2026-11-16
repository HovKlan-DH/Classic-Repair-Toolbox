using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SubmissionCredit - the credit for a draft's work is a row of "Board: Credits" (owner request, 2026-11-15: "as soon as the user
// does a modification for an existing board, then it will populate a new entry in the "Board: Credits", so we will make sure to
// catch this, and so the contributor does not need to manually add a row"). Made with the first change worth a credit, following
// the work while nobody retyped what it says, never beside a credit of the contributor's own, and nothing for a new board.
// ###########################################################################################
public sealed class SubmissionCreditTests
{
    private static ComponentEntry Component(string label, string friendlyName) =>
        new() { BoardLabel = label, FriendlyName = friendlyName, Category = "IC" };

    private static readonly CreditEntry PublishedCredit =
        new() { Category = "Schematic image", SubCategory = "Top", NameOrHandle = "Bo" };

    private static BoardData Published() => new()
    {
        Components = [Component("U8", "CPU")],
        Credits = [PublishedCredit],
    };

    // The published board with its CPU renamed - work CreditSuggestions calls "Board data".
    private static BoardData Renamed(BoardData board, string friendlyName = "ULA") =>
        board.WithComponentSections(components: [Component("U8", friendlyName)]);

    private static BoardData WithPins(BoardData board) =>
        board.WithComponentSections(componentPins: [new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VCC" }]);

    private static BoardData Follow(BoardData published, BoardData before, BoardData after) =>
        SubmissionCredit.Follow(published, before, after, "Anna", "RHD");

    // ---------------------------------------------------------------- Offer

    // Every kind of work once, in the order found - "What" its categories, "Which part of it" its parts - under the name and
    // contact remembered, with no main contribution and no date: the contributor writes the one, the server sets the other.
    [Fact]
    public void The_credit_offered_names_every_kind_of_work_once_under_the_name_remembered()
    {
        CreditSuggestion[] work = [new("Schematic image", "Top"), new("Board labelling", ""), new("Schematic image", "Bottom"), new("board data", ""), new("Board data", "")];

        CreditEntry credit = SubmissionCredit.Offer(work, " Anna ", " RHD ");

        Assert.Equal("Schematic image, Board labelling, board data", credit.Category);
        Assert.Equal("Top, Bottom", credit.SubCategory);
        Assert.Equal(("Anna", "RHD", string.Empty, string.Empty), (credit.NameOrHandle, credit.Contact, credit.MainContribution, credit.Date));
    }

    // ###########################################################################################
    // A name or contact remembered longer than its limit is left out of the credit offered, never cut (code review, 2026-11-16):
    // one remembered before the limits made the credit CRT added an error the contributor never typed - and a cut contact is
    // another address. Left out, the name is asked for as any empty one is.
    // ###########################################################################################
    [Fact]
    public void A_remembered_name_or_contact_past_its_limit_is_left_out()
    {
        CreditSuggestion[] work = [new(CreditSuggestions.BoardData, "")];

        CreditEntry credit = SubmissionCredit.Offer(work, new string('A', CreditLimits.NameOrHandle + 1), "https://example.com/" + new string('x', CreditLimits.Contact));
        Assert.Equal((string.Empty, string.Empty), (credit.NameOrHandle, credit.Contact));

        CreditEntry fits = SubmissionCredit.Offer(work, new string('A', CreditLimits.NameOrHandle), new string('c', CreditLimits.Contact));
        Assert.Equal((CreditLimits.NameOrHandle, CreditLimits.Contact), (fits.NameOrHandle.Length, fits.Contact.Length));
    }

    // "What" is one short line (owner request, 2026-11-15): kinds of work too many for it are "Board data", which covers them.
    [Fact]
    public void Kinds_of_work_too_many_for_a_line_are_Board_data()
    {
        CreditSuggestion[] work =
        [
            new("Schematic image", ""), new("Board labelling", ""), new("Oscilloscope baseline measurement", ""), new("Component pinouts", ""), new("Board Data", ""),
        ];

        CreditEntry credit = SubmissionCredit.Offer(work, "Anna", "");

        Assert.Equal("Board Data", credit.Category);
        Assert.Equal(CreditSuggestions.BoardData, SubmissionCredit.Offer(work[..4], "Anna", "").Category);
        Assert.True(string.Join(", ", work.Select(item => item.Category)).Length > CreditLimits.Category);
    }

    // So is "Which part of it": the parts that fit, and ", ..." for the rest.
    [Fact]
    public void Parts_too_many_for_a_line_end_in_an_ellipsis()
    {
        CreditSuggestion[] work = [new("Schematic image", string.Join(", ", Enumerable.Range(1, 12).Select(index => $"Sheet number {index}")))];

        string parts = SubmissionCredit.Offer(work, "Anna", "").SubCategory;

        Assert.True(parts.Length <= CreditLimits.SubCategory, parts);
        Assert.StartsWith("Sheet number 1, Sheet number 2, ", parts, StringComparison.Ordinal);
        Assert.EndsWith(", ...", parts, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- Follow

    [Fact]
    public void The_first_change_worth_a_credit_adds_one_credit_for_it()
    {
        BoardData published = Published();
        BoardData after = Renamed(published);

        BoardData followed = Follow(published, published, after);

        Assert.Equal(2, followed.Credits.Count);
        CreditEntry credit = followed.Credits[^1];
        Assert.Equal((CreditSuggestions.BoardData, string.Empty, "Anna", "RHD", string.Empty), (credit.Category, credit.SubCategory, credit.NameOrHandle, credit.Contact, credit.MainContribution));
        Assert.Equal(["ULA"], followed.Components.Select(component => component.FriendlyName));
    }

    // Nothing to credit, or no published board to judge by (a new board, whose starting rows are its credits): the board as it is.
    [Fact]
    public void A_change_of_the_credits_alone_or_a_new_board_adds_nothing()
    {
        BoardData published = Published();
        BoardData creditsOnly = published.WithBoardSections(credits: [new CreditEntry { Category = "Schematic image", SubCategory = "Top", NameOrHandle = "Bo B." }]);

        Assert.Same(creditsOnly, Follow(published, published, creditsOnly));

        BoardData after = Renamed(published);
        Assert.Same(after, SubmissionCredit.Follow(null, published, after, "Anna", "RHD"));
    }

    // While its "What" and "Which part of it" are as the rule wrote them, they follow the work - and stay once retyped.
    [Fact]
    public void The_credit_follows_the_work_until_its_What_is_retyped()
    {
        BoardData published = Published();
        BoardData first = Follow(published, published, Renamed(published));
        BoardData second = Follow(published, first, WithPins(first));

        Assert.Equal(2, second.Credits.Count);
        Assert.Equal("Component pinouts, Board data", second.Credits[^1].Category);

        BoardData retyped = second.WithBoardSections(credits: [PublishedCredit, new CreditEntry { Category = "Repairs", NameOrHandle = "Anna", Contact = "RHD" }]);
        BoardData edited = Renamed(retyped, "PLA");

        Assert.Same(edited, Follow(published, retyped, edited));
        Assert.Equal("Repairs", edited.Credits[^1].Category);
    }

    // A credit of the contributor's own is a credit sent: none is added beside it.
    [Fact]
    public void A_credit_of_the_contributors_own_gets_none_beside_it()
    {
        BoardData published = Published();
        BoardData own = Renamed(published).WithBoardSections(credits: [PublishedCredit, new CreditEntry { Category = "Repairs", NameOrHandle = "Anna", MainContribution = "Fixed U8" }]);

        BoardData followed = Follow(published, published, own);

        Assert.Same(own, followed);
    }

    // The edit that took the credit away adds none back - the next change does, as the work needs one.
    [Fact]
    public void A_credit_taken_away_comes_back_with_the_next_change_not_the_same_one()
    {
        BoardData published = Published();
        BoardData credited = Follow(published, published, Renamed(published));
        BoardData removed = credited.WithBoardSections(credits: [PublishedCredit]);

        Assert.Same(removed, Follow(published, credited, removed));

        BoardData next = Follow(published, removed, WithPins(removed));
        Assert.Equal(2, next.Credits.Count);
    }

    // No work left: a credit nothing was written in goes with it; one with a main contribution stays.
    [Fact]
    public void A_credit_nothing_was_written_in_goes_when_no_work_is_left()
    {
        BoardData published = Published();
        BoardData credited = Follow(published, published, Renamed(published));

        BoardData undone = Follow(published, credited, Renamed(credited, "CPU"));
        Assert.Equal([PublishedCredit.Category], undone.Credits.Select(credit => credit.Category));

        BoardData written = credited.WithBoardSections(credits: [PublishedCredit, new CreditEntry { Category = CreditSuggestions.BoardData, NameOrHandle = "Anna", Contact = "RHD", MainContribution = "Renamed U8" }]);
        BoardData kept = Follow(published, written, Renamed(written, "CPU"));
        Assert.Equal(2, kept.Credits.Count);
    }

    // ###########################################################################################
    // *** THE DRAFT IS JUDGED AGAINST THE BOARD AS PUBLISHED NOW (code reviews, 2026-11-16). *** The contributor keeps working
    // in the draft after its first submission - a rename, credited - is published; the draft's base is still the board before
    // it. Judged against the base, the published credit still counted as sent: the next work got no credit of its own, and the
    // published credit's "What" was rewritten. And with only the credits judged against the board now, the published rename
    // still counted as work: every later edit added a credit again for it. Judged against the board as published now (as
    // Submit and the server judge a submission), the next work gets a credit of its own, and an edit adding no work adds none.
    // ###########################################################################################
    [Fact]
    public void A_draft_whose_work_was_published_credits_only_the_work_since()
    {
        BoardData baseBoard = Published();
        var sentFirst = new CreditEntry { Category = CreditSuggestions.BoardData, NameOrHandle = "Anna", Contact = "RHD", MainContribution = "Renamed U8" };
        BoardData credited = Renamed(baseBoard).WithBoardSections(credits: [PublishedCredit, sentFirst]);
        BoardData publishedNow = Renamed(baseBoard).WithBoardSections(credits: [PublishedCredit, sentFirst.WithDate("2026-November-16")]);

        BoardData followed = SubmissionCredit.Follow(publishedNow, credited, WithPins(credited), "Anna", "RHD");

        // The published credit as it is - its "What" not rewritten - and a new one for the pins alone, to be filled in.
        Assert.Equal(3, followed.Credits.Count);
        Assert.Equal((CreditSuggestions.BoardData, "Renamed U8"), (followed.Credits[1].Category, followed.Credits[1].MainContribution));
        Assert.Equal(("Component pinouts", string.Empty), (followed.Credits[^1].Category, followed.Credits[^1].MainContribution));

        // Before Submit too: the work since gets its credit back.
        Assert.Equal(3, SubmissionCredit.Ensure(publishedNow, WithPins(credited), "Anna", "RHD").Credits.Count);

        // An edit adding no work - the published credit's contact corrected, as a maintainer did - adds no credit for the
        // published rename; nor does Submit.
        BoardData contactFixed = credited.WithBoardSections(credits: [PublishedCredit, new CreditEntry { Category = sentFirst.Category, NameOrHandle = "Anna", Contact = "rhd.dk", MainContribution = "Renamed U8" }]);
        Assert.Same(contactFixed, SubmissionCredit.Follow(publishedNow, credited, contactFixed, "Anna", "RHD"));
        Assert.Same(contactFixed, SubmissionCredit.Ensure(publishedNow, contactFixed, "Anna", "RHD"));

        // Judged against the base, as before, the published credit was taken as the one sent.
        Assert.Equal(2, SubmissionCredit.Follow(baseBoard, credited, WithPins(credited), "Anna", "RHD").Credits.Count);
    }

    // ---------------------------------------------------------------- Ensure

    [Fact]
    public void Before_Submit_a_credit_taken_away_is_put_back_and_one_there_is_left_alone()
    {
        BoardData published = Published();
        BoardData uncredited = Renamed(published);

        BoardData ensured = SubmissionCredit.Ensure(published, uncredited, "Anna", "RHD");
        Assert.Equal(2, ensured.Credits.Count);

        Assert.Same(ensured, SubmissionCredit.Ensure(published, ensured, "Anna", "RHD"));
        Assert.Same(published, SubmissionCredit.Ensure(published, published, "Anna", "RHD"));
        Assert.Same(uncredited, SubmissionCredit.Ensure(null, uncredited, "Anna", "RHD"));
    }

    // ---------------------------------------------------------------- what is sent

    // ###########################################################################################
    // *** Which credits a Submit sends is told only against the published board as read, and who is credited is remembered
    // once the send is accepted (code reviews, 2026-11-16). *** With the published board unreadable every credit of the board
    // counted as sent: the summary the maintainers read was every earlier contributor's words, and another person's name became
    // the next draft's credit. And the name was remembered as the dialog opened - a cancelled send too.
    // ###########################################################################################
    [Fact]
    public void Only_the_credits_known_sent_make_the_summary_and_only_an_accepted_send_is_remembered()
    {
        var ray = new CreditEntry { Category = "Board data", NameOrHandle = "Ray", Contact = "his site", MainContribution = "Drew the PCB" };
        var anna = new CreditEntry { Category = "Board data", NameOrHandle = "Anna", Contact = "RHD", MainContribution = "Renamed U8" };
        var board = new BoardData { Credits = [ray, anna] };
        var published = new BoardData { Credits = [ray] };

        IReadOnlyList<CreditEntry> sent = SubmissionCredit.KnownSent(isNewBoard: false, board, published);
        Assert.Equal([anna], sent);
        Assert.Equal("Renamed U8", SubmissionCredit.MainContributionsOf(sent));

        // The published board not read: which were sent is not known - Ray's words are not the summary, nor his name remembered.
        Assert.Empty(SubmissionCredit.KnownSent(isNewBoard: false, board, published: null));

        // A new board's credits are all its own.
        Assert.Equal([ray, anna], SubmissionCredit.KnownSent(isNewBoard: true, board, published: null));

        // Remembered once accepted - not cancelled or failed (no result), nor refused.
        var accepted = new SubmissionResult { IsAccepted = true };
        Assert.Equal([anna], SubmissionCredit.ToRemember(accepted, isNewBoard: false, sent));
        Assert.Empty(SubmissionCredit.ToRemember(null, isNewBoard: false, sent));
        Assert.Empty(SubmissionCredit.ToRemember(new SubmissionResult { IsAccepted = false }, isNewBoard: false, sent));
    }

    // ###########################################################################################
    // *** Only credits naming ONE person teach CRT who the contributor is (code review, 2026-11-16). *** A new board's starting
    // rows credit whoever drew each schematic image, and a credit added for somebody else could stand first: remembered, that
    // person became the next draft's credit. "Unknown", nobody, does not count as a person.
    // ###########################################################################################
    [Fact]
    public void Only_credits_naming_one_person_are_remembered_and_never_a_new_boards()
    {
        var ray = new CreditEntry { Category = "Board data", NameOrHandle = "Ray", Contact = "his site" };
        var anna = new CreditEntry { Category = "Board data", NameOrHandle = "Anna", Contact = "RHD", MainContribution = "Renamed U8" };
        var annaAgain = new CreditEntry { Category = "Schematic image", NameOrHandle = "anna", Contact = "rhd", MainContribution = "Drew Top" };
        var nobody = new CreditEntry { Category = "Board data", NameOrHandle = "Unknown" };
        var accepted = new SubmissionResult { IsAccepted = true };

        Assert.Empty(SubmissionCredit.ToRemember(accepted, isNewBoard: true, [anna]));
        Assert.Empty(SubmissionCredit.ToRemember(accepted, isNewBoard: false, [ray, anna]));
        Assert.Equal([anna, annaAgain], SubmissionCredit.ToRemember(accepted, isNewBoard: false, [anna, nobody, annaAgain]));
    }

    // The credits sent are those the published board does not have - their main contributions what the submission says it changed.
    [Fact]
    public void The_credits_sent_are_those_the_published_board_has_not()
    {
        BoardData published = Published();
        var mine = new CreditEntry { Category = CreditSuggestions.BoardData, NameOrHandle = "Anna", MainContribution = "Renamed **U8**" };
        BoardData draft = Renamed(published).WithBoardSections(credits: [PublishedCredit, mine]);

        Assert.Equal(["Anna"], SubmissionCredit.SentWith(draft, published).Select(credit => credit.NameOrHandle));
        Assert.Equal("Renamed **U8**", SubmissionCredit.MainContributionsOf(SubmissionCredit.SentWith(draft, published)));

        // A new board's every credit is sent.
        Assert.Equal(2, SubmissionCredit.SentWith(draft, null).Count);
    }
}
