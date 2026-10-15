using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ReviewApiParser.ParseSubmission - the answer a maintainer opens a submission on.
//
// The shapes here were verified by SERIALISING a real ReviewChangeSummary through
// System.Text.Json with ASP.NET Core's camelCase policy, not by reading the record declarations.
// That mattered: the records carry computed properties (`hasChanges`, `totalChanges`) which also
// appear on the wire, and the server sends all ten sections every time including empty ones -
// both facts a guess would have missed.
//
// ReviewWireContractTests runs a genuine answer through serialise-then-parse; this file covers the
// degenerate answers that a real one never produces.
public sealed class ReviewSubmissionDetailTests
{
    private const string DetailJson = """
        {"canPublish":true,
         "submission":{"id":42,"boardId":"Commodore/C64/250407","state":"pending",
                       "summary":"Corrected R12.","contactEmail":"someone@example.com",
                       "createdUtc":"2026-09-21T12:00:00+00:00"},
         "findings":[],
         "changes":{"isNewBoard":false,"revisionDate":null,"sections":[
            {"section":"Board schematics","added":[],"removed":[],"changed":[],"renamed":[],
             "hasChanges":false,"totalChanges":0},
            {"section":"Components","added":["U9"],"removed":[],"changed":["U8"],"renamed":[],
             "hasChanges":true,"totalChanges":2}]}}
        """;

    [Fact]
    public void A_detail_response_yields_the_submission_and_its_changes()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(ReviewSubmissionDetailTests.DetailJson);

        Assert.NotNull(detail);
        Assert.True(detail!.CanPublish);
        Assert.Equal(42, detail.Submission.Id);
        Assert.Equal("Corrected R12.", detail.Submission.Summary);

        Assert.NotNull(detail.Changes);
        Assert.False(detail.Changes!.IsNewBoard);
        Assert.Equal(2, detail.Changes.TotalChanges);
    }

    [Fact]
    public void EMPTY_sections_are_dropped_at_the_edge()
    {
        // The server sends all ten every time (cheap - about 1.4 KB), but the maintainer's screen
        // only ever shows the ones that changed. Filtering here means nothing downstream has to
        // remember to.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(ReviewSubmissionDetailTests.DetailJson);

        ReviewSectionView section = Assert.Single(detail!.Changes!.Sections);

        Assert.Equal("Components", section.Section);
        Assert.Equal(["U9"], section.Added);
        Assert.Equal(["U8"], section.Changed);
    }

    [Fact]
    public void A_NULL_changes_object_is_not_the_same_as_NO_changes()
    {
        // *** THE DISTINCTION THAT PROTECTS A MAINTAINER. *** The server answers changes:null when
        // a submission's payload could not be loaded. Rendering that as "no changes" would invite
        // somebody to approve a submission nobody has been able to look at.
        ReviewSubmissionDetail? unloadable = ReviewApiParser.ParseSubmission("""
            {"canPublish":false,"submission":{"id":1,"boardId":"A/B/C","state":"pending"},
             "changes":null,"findings":[]}
            """);

        Assert.NotNull(unloadable);
        Assert.Null(unloadable!.Changes);

        ReviewSubmissionDetail? unchanged = ReviewApiParser.ParseSubmission("""
            {"canPublish":false,"submission":{"id":1,"boardId":"A/B/C","state":"pending"},
             "changes":{"isNewBoard":false,"sections":[]},"findings":[]}
            """);

        Assert.NotNull(unchanged!.Changes);
        Assert.False(unchanged.Changes!.HasChanges);
    }

    [Fact]
    public void A_submission_with_no_summary_still_opens_so_its_findings_can_be_read()
    {
        // The reason a submission has no comparable payload is in its findings. Refusing the
        // whole response would hide the very explanation the maintainer came for.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"canPublish":false,"submission":{"id":1,"boardId":"A/B/C","state":"rejected"},
             "changes":null,
             "findings":[{"severity":1,"code":"file.missing","subject":"Images/u8.png",
                          "message":"The file is referenced but was not uploaded."}]}
            """);

        Assert.NotNull(detail);

        ReviewFindingView finding = Assert.Single(detail!.Findings);
        Assert.True(finding.IsError);
        Assert.Equal("file.missing", finding.Code);
        Assert.Equal("Images/u8.png", finding.Subject);
    }

    [Fact]
    public void Severity_is_read_as_a_NUMBER_or_a_STRING()
    {
        // ValidationSeverity serialises as a number by default (Warning = 0, Error = 1). If the
        // server ever adds a JsonStringEnumConverter, a number-only reader would silently turn
        // every error into a warning - on the screen whose job is to say what is wrong.
        ReviewSubmissionDetail? numeric = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[{"severity":1,"message":"a"},{"severity":0,"message":"b"}]}
            """);

        Assert.True(numeric!.Findings[0].IsError);
        Assert.False(numeric.Findings[1].IsError);

        ReviewSubmissionDetail? textual = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[{"severity":"Error","message":"a"},{"severity":"Warning","message":"b"}]}
            """);

        Assert.True(textual!.Findings[0].IsError);
        Assert.False(textual.Findings[1].IsError);
    }

    [Fact]
    public void A_finding_with_no_message_is_KEPT_rather_than_dropped()
    {
        // These are the reasons a submission was flagged. Dropping one because its prose is
        // missing would leave a maintainer looking at a shorter list than the truth.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[{"severity":1,"code":"scale.components"}]}
            """);

        ReviewFindingView finding = Assert.Single(detail!.Findings);

        Assert.Equal("scale.components", finding.Code);
        Assert.Equal(string.Empty, finding.Message);
    }

    // ###########################################################################################
    // *** THE PER-FILE FACTS SURVIVE THE WIRE, UNDER THE SERVER'S OWN JSON SETTINGS (security
    // review, 2026-09-25). *** CRT.Server serialises CRT.Data's SubmittedFileFact with camelCase
    // names and nulls omitted (Program.cs, ConfigureHttpJsonOptions); the Maintainer tab reads the same type
    // back. Serialising it here with those settings and parsing the result is the closest a test
    // in this project can come to both ends at once: a renamed property, or a scope that travelled
    // as a number, fails here rather than as a blank list on a maintainer's screen.
    // ###########################################################################################
    [Fact]
    public void Submitted_file_facts_round_trip_under_the_servers_json_settings()
    {
        var serverOptions = new System.Text.Json.JsonSerializerOptions
        {
            PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        SubmittedFileFact[] sent =
        [
            new("Commodore/C64/250407/manual.pdf", new string('1', 64), 12, SubmissionFileScope.Own, true, null),
            new("Commodore/Shared files/x.png", new string('2', 64), 34, SubmissionFileScope.ManufacturerShared, false, new string('3', 64))
        ];

        string json = System.Text.Json.JsonSerializer.Serialize(
            new { submission = new { id = 1, boardId = "Commodore/C64/250407", state = "pending" }, submittedFiles = sent },
            serverOptions);

        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(json);

        Assert.NotNull(detail);
        Assert.Equal(sent, detail!.SubmittedFiles);
    }

    // An older server that does not send the field, and a malformed entry, both degrade quietly:
    // no list, or the list without that one entry - never a failure to open the submission.
    [Fact]
    public void Missing_or_malformed_file_facts_never_stop_the_submission_opening()
    {
        ReviewSubmissionDetail? absent = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[]}
            """);

        ReviewSubmissionDetail? partly = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "submittedFiles":[
               {"path":"A/B/C/a.png","sha256":"aa","sizeBytes":1,"scope":"Own","isReferenced":true},
               {"path":"A/B/C/b.png","sha256":"bb","sizeBytes":1,"scope":"NoSuchScope","isReferenced":true},
               {"path":"","sha256":"cc","sizeBytes":1,"scope":"Own","isReferenced":true}
             ]}
            """);

        Assert.Empty(absent!.SubmittedFiles);
        Assert.Equal("A/B/C/a.png", Assert.Single(partly!.SubmittedFiles).Path);
    }

    [Fact]
    public void A_rename_survives_the_wire()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "changes":{"isNewBoard":false,"sections":[
                {"section":"Components","added":[],"removed":[],"changed":[],
                 "renamed":[{"from":"U8","to":"U9","alsoChanged":true}],
                 "hasChanges":true,"totalChanges":1}]}}
            """);

        ReviewRenameView renamed = Assert.Single(Assert.Single(detail!.Changes!.Sections).Renamed);

        Assert.Equal("U8", renamed.From);
        Assert.Equal("U9", renamed.To);
        Assert.True(renamed.AlsoChanged);
    }

    [Fact]
    public void A_half_written_rename_is_skipped_rather_than_drawn_as_an_arrow_to_nowhere()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "changes":{"isNewBoard":false,"sections":[
                {"section":"Components","added":[],"removed":[],"changed":[],
                 "renamed":[{"from":"U8"},{"to":"U9"},{"from":"A","to":"B"}],
                 "hasChanges":true,"totalChanges":3}]}}
            """);

        ReviewRenameView renamed = Assert.Single(Assert.Single(detail!.Changes!.Sections).Renamed);

        Assert.Equal("A", renamed.From);
    }

    [Fact]
    public void A_new_board_survives_the_wire()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "changes":{"isNewBoard":true,"sections":[
                {"section":"Components","added":["U1","U2"],"removed":[],"changed":[],"renamed":[],
                 "hasChanges":true,"totalChanges":2}]}}
            """);

        Assert.True(detail!.Changes!.IsNewBoard);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("""{"canPublish":true}""")]
    public void An_answer_that_is_not_a_submission_yields_null_rather_than_throwing(string? body)
    {
        Assert.Null(ReviewApiParser.ParseSubmission(body));
    }

    [Fact]
    public void A_submission_with_no_id_is_refused()
    {
        // Without an id nothing further can be requested about it, so it is not a usable answer.
        Assert.Null(ReviewApiParser.ParseSubmission("""{"submission":{"boardId":"A/B/C"}}"""));
    }

    [Fact]
    public void A_response_with_NO_manifest_at_all_still_parses()
    {
        // The server answers with a null manifest when a payload could not be loaded. The
        // maintainer sees the findings explaining why; the screen must still draw.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[]}
            """);

        Assert.NotNull(detail);
        Assert.Empty(detail!.SubmittedFiles);
    }

    // ###########################################################################################
    // *** THE FIELDS ONLY THE RETIRED CHANGE SUMMARY READ ARE READ PAST (code review, 2026-10-13). ***
    // The server still sends manifest, publishedFiles, publishedHashes and schematicImages -
    // released CRT 3.0.0 reads them - but this CRT reads none of them since the summary that drew
    // them was deleted. An answer carrying them parses as one without them.
    // ###########################################################################################
    [Fact]
    public void The_fields_only_the_retired_change_summary_read_are_read_past()
    {
        const string bare = """{"canPublish":true,"submission":{"id":7,"boardId":"Commodore/C64/250407"},"findings":[]}""";

        const string carrying = """
            {"canPublish":true,"submission":{"id":7,"boardId":"Commodore/C64/250407"},"findings":[],
             "manifest":{"files":[{"path":"Images/U8.png","sha256":"aaaa"}]},
             "publishedFiles":["Images/U8.png"],
             "publishedHashes":{"Images/U8.png":"bbbb"},
             "schematicImages":{"Sheet 1":"Images/sheet1.png"}}
            """;

        ReviewSubmissionDetail? withThem = ReviewApiParser.ParseSubmission(carrying);
        ReviewSubmissionDetail? without = ReviewApiParser.ParseSubmission(bare);

        Assert.NotNull(withThem);
        Assert.Equal(7, withThem!.Submission.Id);
        Assert.True(withThem.CanPublish);
        Assert.Equal(without!.Submission, withThem.Submission);
        Assert.Empty(withThem.SubmittedFiles);
    }
}
