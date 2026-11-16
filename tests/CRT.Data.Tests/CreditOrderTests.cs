using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A credit's "Main contribution" and "Date" (owner request, 2026-11-14: "two more columns named "Main contribution" and
// "Date" ... have the credits now sorted by oldest contribution first"; owner decisions: the main contribution required on
// "New credits only", the date "Set by the server at publish", credits "Always kept in date order"; and 2026-11-15: "the
// existing credits should actually have NO date at all - but all new credits must have it set"). The rules are CRT.Data's
// and both sides go by them: CreditOrder, BoardRowRequired with PublishedCredits, the checks, the differ, the table's cells,
// PublishMerge and the board file.
// ###########################################################################################
public sealed class CreditOrderTests
{
    private const string Board = "Commodore/C64/250407";

    private static readonly DateTimeOffset December3 = new(2026, 12, 3, 22, 15, 0, TimeSpan.Zero);

    private static CreditEntry Credit(string name, string date = "", string contribution = "", string category = "Board data") =>
        new() { Category = category, NameOrHandle = name, Date = date, MainContribution = contribution };

    // ------------------------------------------------------------------ The order

    // The undated first - from before the history was kept, so older than any dated one - in the board's order; then the
    // dated by day, never by the text, where "August" comes before "May"; a date that cannot be read last.
    [Fact]
    public void Credits_go_undated_first_then_oldest_first_and_the_unreadable_last()
    {
        List<CreditEntry> sorted = CreditOrder.Sorted(
        [
            Credit("Unreadable", "someday"),
            Credit("Old one"),
            Credit("August", "2027-August-3"),
            Credit("May first", "2027-May-12"),
            Credit("Old two"),
            Credit("May second", "2027-May-12"),
        ]);

        Assert.Equal(["Old one", "Old two", "May first", "May second", "August", "Unreadable"], sorted.Select(credit => credit.NameOrHandle));
    }

    // ###########################################################################################
    // A publish dates the credits NEW in it, the day it is published - and only those: a credit the board has published
    // keeps its date, and one from before the history stays undated (owner decision, 2026-11-15).
    //
    // *** A NEW CREDIT CANNOT BE BACK-DATED (owner decision, 2026-11-16: "Should not be possible to back-date - new
    // contributions, not seen before, must be set when board maintainer pushes this to BETA and STABLE"). *** The date a new
    // credit was sent with is let go of - it was kept until then ("Given" stood dated May, before every December credit) -
    // while a credit the board has, sent with another date, keeps it: a maintainer's correction.
    // ###########################################################################################
    [Fact]
    public void A_publish_dates_only_the_credits_new_in_it_and_never_as_they_were_sent()
    {
        List<CreditEntry> dated = CreditOrder.Dated(
            [Credit("New", contribution: "Drew the PCB"), Credit("anna "), Credit("Cy"), Credit("Given", "2026-May-12", "Added it")],
            [Credit("Anna"), Credit("Cy", "2026-November-20", "Fixed U8")],
            December3);

        Assert.Equal(["anna ", "Cy", "New", "Given"], dated.Select(credit => credit.NameOrHandle));
        Assert.Equal([string.Empty, "2026-November-20", "2026-December-3", "2026-December-3"], dated.Select(credit => credit.Date));

        // The credit the board has, its date corrected: the correction stands.
        Assert.Equal(
            "2026-November-21",
            Assert.Single(CreditOrder.Dated([Credit("Cy", "2026-November-21", "Fixed U8")], [Credit("Cy", "2026-November-20", "Fixed U8")], December3)).Date);
    }

    // ###########################################################################################
    // *** A schematic image's credit following its image's rename is still the published credit (code review, 2026-11-16). ***
    // A picture renamed on its card takes the credits for it along (SchematicRename): the author's credit then read as changed
    // - a main contribution demanded of somebody else's credit, Submit stopped, the credit sent and dated the day of the
    // publish. The check and the dating agree it is the one published; a credit of another kind is not matched so.
    // ###########################################################################################
    [Fact]
    public void A_schematic_image_credit_that_followed_its_images_rename_stands_as_published()
    {
        var published = new CreditEntry { Category = CreditSuggestions.SchematicImage, SubCategory = "Top", NameOrHandle = "Thomas", Contact = "his site" };
        var followed = new CreditEntry { Category = CreditSuggestions.SchematicImage, SubCategory = "PCB Top", NameOrHandle = "Thomas", Contact = "his site" };

        Assert.Empty(PublishedCredits.Judged([followed], PublishedCredits.Of([published])));
        Assert.Empty(CreditOrder.NotPublished([followed], [published]));
        Assert.Equal(string.Empty, Assert.Single(CreditOrder.Dated([followed], [published], December3)).Date);

        // Each published credit answers for ONE: the image's own credit still there, a second of the person for another image is new.
        var another = new CreditEntry { Category = CreditSuggestions.SchematicImage, SubCategory = "Bottom", NameOrHandle = "Thomas", Contact = "his site" };
        Assert.Equal([another], PublishedCredits.Judged([published, another], PublishedCredits.Of([published])));

        // Another kind of credit with another part is another credit.
        var boardData = new CreditEntry { Category = "Board data", SubCategory = "PAL", NameOrHandle = "Thomas" };
        Assert.Single(PublishedCredits.Judged([boardData], PublishedCredits.Of([new CreditEntry { Category = "Board data", SubCategory = "NTSC", NameOrHandle = "Thomas" }])));
    }

    // A date given on a credit the board has is published trimmed, as every reading of it trims (code review, 2026-11-16:
    // written with its spaces).
    [Fact]
    public void A_date_given_is_published_trimmed()
    {
        List<CreditEntry> dated = CreditOrder.Dated([Credit("Given", " 2026-May-12 ", "Added it")], [Credit("Given", "2026-May-1", "Added it")], December3);

        Assert.Equal("2026-May-12", Assert.Single(dated).Date);
    }

    // ###########################################################################################
    // EACH PUBLISHED CREDIT ANSWERS FOR ONE CREDIT SENT. The same person credited again for the same kind of work is a new
    // line of the history - its very values (a main contribution of its own) find no published credit, and the one with the
    // same name and kind is the published one's already. So it is dated, and the old one keeps having none.
    // ###########################################################################################
    [Fact]
    public void The_same_person_credited_again_for_the_same_work_is_a_new_credit()
    {
        IReadOnlyList<CreditEntry> published = [Credit("Bo")];
        List<CreditEntry> sent = [Credit("Bo"), Credit("Bo", contribution: "Added the pinouts")];

        List<CreditEntry> dated = CreditOrder.Dated(sent, published, December3);

        Assert.Equal([string.Empty, "2026-December-3"], dated.Select(credit => credit.Date));
        Assert.Equal([string.Empty, "Added the pinouts"], dated.Select(credit => credit.MainContribution));

        // The pane says the same of the draft: only the second is "Dated when published".
        CreditEntry added = Assert.Single(CreditOrder.NotPublished(sent, published));
        Assert.Same(sent[1], added);
    }

    // A credit changed (another contact) is still the published one: it keeps its date, or its having none.
    [Fact]
    public void A_credit_changed_keeps_its_published_date()
    {
        List<CreditEntry> dated = CreditOrder.Dated(
            [new CreditEntry { Category = "Board data", NameOrHandle = "Cy", Contact = "a forum", MainContribution = "Fixed U8" }],
            [Credit("Cy", "2026-November-20", "Fixed U8")],
            December3);

        Assert.Equal("2026-November-20", Assert.Single(dated).Date);
        Assert.Empty(CreditOrder.NotPublished([Credit("Anna", contribution: "Drew it")], [Credit("Anna")]));
    }

    // ###########################################################################################
    // A credit sent undated that the board has published keeps its published date at the review (Build): the contributor's
    // own credit, undated in the draft once it was published, must not become the day of the next publish. A credit the
    // board does not have is undated, for the publish to date - also one sent with a date (owner decision, 2026-11-16: "Should
    // not be possible to back-date"), so the review never shows one.
    // ###########################################################################################
    [Fact]
    public void A_credit_sent_undated_that_the_board_has_keeps_its_published_date()
    {
        List<CreditEntry> merged = CreditOrder.WithPublishedDates(
            [Credit("anna "), Credit("Bo"), Credit("Cy"), Credit("Dee", "2019-January-1")],
            [Credit("Anna", "2026-November-20"), Credit("Cy")]);

        Assert.Equal(["2026-November-20", string.Empty, string.Empty, string.Empty], merged.Select(credit => credit.Date));
    }

    [Fact]
    public void Build_carries_the_published_dates_and_the_publish_dates_the_new_ones()
    {
        var manifest = new SubmissionManifest
        {
            BoardId = Board,
            Rows = new SubmissionRows { Credits = [Credit("Bo", contribution: "Drew the PCB"), Credit("Anna"), Credit("Old")] },
        };

        BoardData published = new() { Credits = [Credit("Old"), Credit("Anna", "2026-November-20")] };

        BoardData board = PublishMerge.WithCreditsDated(PublishMerge.Build(manifest, published), published, December3);

        Assert.Equal(["Old", "Anna", "Bo"], board.Credits.Select(credit => credit.NameOrHandle));
        Assert.Equal([string.Empty, "2026-November-20", "2026-December-3"], board.Credits.Select(credit => credit.Date));
        Assert.Equal("Drew the PCB", board.Credits[2].MainContribution);
    }

    // A new board's every credit is new: each is dated.
    [Fact]
    public void A_new_board_s_credits_are_all_dated()
    {
        List<CreditEntry> dated = CreditOrder.Dated([Credit("Anna", contribution: "Drew it"), Credit("Bo", contribution: "Measured it")], published: null, December3);

        Assert.All(dated, credit => Assert.Equal("2026-December-3", credit.Date));
    }

    // "Board: Credits" writes them in the board's order, a credit added standing last: only a publish can tell an old
    // undated credit from a new one.
    [Fact]
    public void Board_credits_are_written_in_the_board_s_order()
    {
        BoardData board = BoardPartEdits.WithCredits(
            new BoardData(),
            [Credit("Dated", "2026-December-3", "Drew it"), Credit("Old"), Credit("New", contribution: "Added the pinouts")]);

        Assert.Equal(["Dated", "Old", "New"], board.Credits.Select(credit => credit.NameOrHandle));
    }

    // ------------------------------------------------------------------ The main contribution, required on new credits

    // ###########################################################################################
    // Published as it stands, a credit needs no main contribution - none of those published before 2026-11-14 has one. Any
    // other credit does: one added, or one changed (another contact), a new board's every credit. Null judges none.
    // ###########################################################################################
    [Fact]
    public void Only_a_credit_not_published_as_it_stands_needs_its_main_contribution()
    {
        // A credit alone on its board, judged as the checks judge the list (PublishedCredits.Judged).
        static IReadOnlyList<string> EmptyOf(CreditEntry row, PublishedCredits? published) =>
            BoardRowRequired.EmptyColumns(row, PublishedCredits.Judged([row], published).Contains(row));

        CreditEntry old = Credit("Anna");
        PublishedCredits published = PublishedCredits.Of([old]);

        Assert.Empty(EmptyOf(old, published));
        Assert.Equal([BoardWorkbookSchema.ColMainContribution], EmptyOf(Credit("Bo"), published));
        Assert.Equal([BoardWorkbookSchema.ColMainContribution], EmptyOf(new CreditEntry { Category = "Board data", NameOrHandle = "Anna", Contact = "a forum" }, published));
        Assert.Equal([BoardWorkbookSchema.ColMainContribution], EmptyOf(old, PublishedCredits.None));
        Assert.Empty(EmptyOf(Credit("Bo", contribution: "- Drew the **top** side"), published));
        Assert.Empty(EmptyOf(Credit("Bo"), published: null));
    }

    // The table, the draft's row and Submit say it as the server refuses it - one rule, its error on the very cell.
    [Fact]
    public void A_new_credit_without_its_main_contribution_is_refused_by_the_server_as_the_table_marks_it()
    {
        var manifest = new SubmissionManifest
        {
            BoardId = Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = "A credit.",
            Rows = new SubmissionRows
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" }],
                Credits = [Credit("Anna"), Credit("Bo")],
            },
        };

        PublishedCredits published = PublishedCredits.Of([Credit("Anna")]);

        ValidationFinding refused = Assert.Single(SubmissionValidator.Validate(manifest, [$"{Board}/main.png"], published), finding => finding.Code == "row.required_empty");
        Assert.Equal(BoardWorkbookSchema.ColMainContribution, refused.Subject);

        BoardDataProblem marked = Assert.Single(
            BoardDataChecks.Check(BoardCheckRows.From(manifest.Rows, published), null, BoardCheckScope.Everything),
            problem => problem.Code == "row.required_empty");
        Assert.Equal((BoardWorkbookSchema.SheetCredits, 1, BoardWorkbookSchema.ColMainContribution), (marked.Sheet, marked.Index, marked.Column));

        // At the finalise the rows were judged already: nothing is said of it.
        Assert.DoesNotContain(SubmissionValidator.Validate(manifest, [$"{Board}/main.png"], publishedCredits: null), finding => finding.Code == "row.required_empty");
    }

    // ###########################################################################################
    // *** A PUBLISHED CREDIT STANDS FOR ONE ROW (owner report, 2026-11-15: the credit made for a change came out the same as one
    // the board had published - "Board data", the same name and contact, no main contribution - and went out with its main
    // contribution empty). *** Asked by its values alone, the second row "stood as published" too. Each published credit answers
    // for ONE row, the first in the board's order; another like it is new and needs its main contribution - in the table and the
    // draft's checks, in the readiness count and at the server, BETA's and stable's credits together counting once.
    // ###########################################################################################
    [Fact]
    public void A_second_credit_like_a_published_one_needs_its_main_contribution()
    {
        static CreditEntry Johnny() => new() { Category = "Board data", NameOrHandle = "Johnny boy", Contact = "Nowhere to be found" };

        PublishedCredits published = PublishedCredits.Of([Johnny()]);
        var board = new BoardData { Credits = [Johnny(), Johnny()] };

        foreach (BoardCheckScope scope in new[] { BoardCheckScope.Submission, BoardCheckScope.Everything })
        {
            BoardDataProblem marked = Assert.Single(
                BoardDataChecks.Check(BoardCheckRows.From(board, published), null, scope),
                problem => problem.Code == "row.required_empty");
            Assert.Equal((BoardWorkbookSchema.SheetCredits, 1, BoardWorkbookSchema.ColMainContribution), (marked.Sheet, marked.Index, marked.Column));
        }

        Assert.Equal(1, BoardRowRequired.Of(board, published).Credits);

        var manifest = new SubmissionManifest
        {
            BoardId = Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = "A credit.",
            Rows = new SubmissionRows { Credits = [Johnny(), Johnny()] },
        };

        Assert.Single(SubmissionValidator.Validate(manifest, [], PublishedCredits.Union(published, PublishedCredits.Of([Johnny()]))), finding => finding.Code == "row.required_empty");

        // Published twice, both stand.
        Assert.DoesNotContain(
            BoardDataChecks.Check(BoardCheckRows.From(board, PublishedCredits.Of([Johnny(), Johnny()])), null, BoardCheckScope.Everything),
            problem => problem.Code == "row.required_empty");
    }

    // A date that cannot be read is refused in both scopes - the order would put it out of place for good. None is fine.
    [Fact]
    public void A_credit_date_that_cannot_be_read_is_an_error_on_both_sides()
    {
        BoardCheckRows rows = BoardCheckRows.From(new BoardData { Credits = [Credit("Anna", "last summer", "Drew it")] }, published: null);

        foreach (BoardCheckScope scope in new[] { BoardCheckScope.Submission, BoardCheckScope.Everything })
        {
            BoardDataProblem problem = Assert.Single(BoardDataChecks.Check(rows, null, scope), each => each.Code == "credit.date_unreadable");
            Assert.Equal(BoardProblemLevel.Error, problem.Level);
            Assert.Equal(BoardWorkbookSchema.ColDate, problem.Column);
        }

        foreach (string fine in new[] { "2026-November-14", string.Empty })
        {
            Assert.DoesNotContain(
                BoardDataChecks.Check(BoardCheckRows.From(new BoardData { Credits = [Credit("Anna", fine, "Drew it")] }, published: null), null, BoardCheckScope.Everything),
                problem => problem.Code == "credit.date_unreadable");
        }
    }

    // ------------------------------------------------------------------ An empty date in a draft is no change

    // ###########################################################################################
    // The server dates a credit as it publishes, so the contributor's own credit, undated in the draft, is no change against
    // the published one: neither the differ nor the table counts it - else a draft fully published read "1 row changed" and
    // was never retired. A date given that differs still is a change (a maintainer's correction).
    // ###########################################################################################
    [Fact]
    public void An_undated_credit_in_a_draft_is_no_change_against_its_published_date()
    {
        BoardData published = new() { Credits = [Credit("Bo", "2026-December-3", "Drew the PCB")] };

        Assert.Equal(0, BoardDataDiffer.CountChanges(published, new BoardData { Credits = [Credit("Bo", contribution: "Drew the PCB")] }));
        Assert.Equal(1, BoardDataDiffer.CountChanges(published, new BoardData { Credits = [Credit("Bo", "2026-December-4", "Drew the PCB")] }));
        Assert.Equal(1, BoardDataDiffer.CountChanges(published, new BoardData { Credits = [Credit("Bo", "2026-December-3", "Drew the board")] }));
    }

    [Fact]
    public void The_table_colors_no_undated_credit_date_as_changed()
    {
        BoardData published = new() { Credits = [Credit("Bo", "2026-December-3", "Drew the PCB")] };

        BoardTableDocument document = BoardTableDocument.Create(published, new BoardData { Credits = [Credit("Bo", contribution: "Drew the PCB")] });
        BoardTableSheet credits = document.FindSheet(BoardWorkbookSchema.SheetCredits)!;

        Assert.Equal(0, credits.ChangeCount);
        Assert.True(BoardWorkbookSchema.Credits.IsSetWhenPublished(BoardWorkbookSchema.ColDate));
    }

    // ###########################################################################################
    // Two credits "Unknown" for the same kind of work - two contributors who wanted no name - are two rows, told apart by
    // their main contributions (BoardDraftNaturalKeys.ForCredit): one added beside the other is one row added, never the
    // first changed.
    // ###########################################################################################
    [Fact]
    public void Two_unknown_credits_for_the_same_work_are_two_rows()
    {
        CreditEntry first = new() { Category = "Board data", NameOrHandle = CreditSuggestions.UnknownName, MainContribution = "Fixed U8's part number" };
        CreditEntry second = new() { Category = "Board data", NameOrHandle = CreditSuggestions.UnknownName, MainContribution = "Added the pinouts" };

        BoardData published = new() { Credits = [first] };

        Assert.Equal(1, BoardDataDiffer.CountChanges(published, new BoardData { Credits = [first, second] }));
        Assert.NotEqual(BoardDraftNaturalKeys.ForRow(first), BoardDraftNaturalKeys.ForRow(second));
    }

    // ------------------------------------------------------------------ The board file

    // Written only when filled, so a board without them keeps its bytes; read back as they were.
    [Fact]
    public void The_board_file_carries_the_main_contribution_and_the_date_only_when_filled()
    {
        BoardData board = new() { Credits = [Credit("Anna"), Credit("Bo", "2026-December-3", "Drew **the** PCB")] };

        BoardDocument document = BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, []);

        Assert.Null(document.Credits[0].MainContribution);
        Assert.Null(document.Credits[0].Date);
        Assert.Equal("Drew **the** PCB", document.Credits[1].MainContribution);
        Assert.Equal("2026-December-3", document.Credits[1].Date);

        CreditEntry back = BoardDocumentMapping.ToContent(document).Board.Credits[1];
        Assert.Equal(("Drew **the** PCB", "2026-December-3"), (back.MainContribution, back.Date));
    }
}
