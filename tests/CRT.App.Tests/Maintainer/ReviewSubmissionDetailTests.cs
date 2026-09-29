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
// ReviewSummaryPresenterTests goes further and runs a genuine comparison through serialise-then-
// parse; this file covers the degenerate answers that a real comparison never produces.
public sealed class ReviewSubmissionDetailTests
{
    private const string DetailJson = """
        {"canPublish":true,
         "submission":{"id":42,"systemId":"Commodore/C64/250407","state":"pending",
                       "summary":"Corrected R12.","contactEmail":"someone@example.com",
                       "createdUtc":"2026-09-21T12:00:00+00:00"},
         "findings":[],
         "changes":{"isNewSystem":false,"revisionDate":null,"sections":[
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
        Assert.False(detail.Changes!.IsNewSystem);
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
            {"canPublish":false,"submission":{"id":1,"systemId":"A/B/C","state":"pending"},
             "changes":null,"findings":[]}
            """);

        Assert.NotNull(unloadable);
        Assert.Null(unloadable!.Changes);

        ReviewSubmissionDetail? unchanged = ReviewApiParser.ParseSubmission("""
            {"canPublish":false,"submission":{"id":1,"systemId":"A/B/C","state":"pending"},
             "changes":{"isNewSystem":false,"sections":[]},"findings":[]}
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
            {"canPublish":false,"submission":{"id":1,"systemId":"A/B/C","state":"rejected"},
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
            new { submission = new { id = 1, systemId = "Commodore/C64/250407", state = "pending" }, submittedFiles = sent },
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
             "changes":{"isNewSystem":false,"sections":[
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
             "changes":{"isNewSystem":false,"sections":[
                {"section":"Components","added":[],"removed":[],"changed":[],
                 "renamed":[{"from":"U8"},{"to":"U9"},{"from":"A","to":"B"}],
                 "hasChanges":true,"totalChanges":3}]}}
            """);

        ReviewRenameView renamed = Assert.Single(Assert.Single(detail!.Changes!.Sections).Renamed);

        Assert.Equal("A", renamed.From);
    }

    [Fact]
    public void A_new_system_survives_the_wire()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "changes":{"isNewSystem":true,"sections":[
                {"section":"Components","added":["U1","U2"],"removed":[],"changed":[],"renamed":[],
                 "hasChanges":true,"totalChanges":2}]}}
            """);

        Assert.True(detail!.Changes!.IsNewSystem);
        Assert.Contains("New system", ReviewSummaryPresenter.BuildHeadline(detail.Changes));
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
        Assert.Null(ReviewApiParser.ParseSubmission("""{"submission":{"systemId":"A/B/C"}}"""));
    }

    // -----------------------------------------------------------------------------------------
    // The manifest's file list, which drives task 4's visual comparison.
    //
    // This JSON was verified by SERIALISING a real SubmissionManifest with camelCase, the same
    // practice the header describes - the nesting (manifest.files[].sha256) is not something to
    // guess at, because a wrong key parses to an empty list and the comparison silently shows
    // nothing rather than failing.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_manifests_FILES_are_carried_through_for_fetching()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "manifest":{"systemId":"Commodore/C64/250407","files":[
                {"path":"Images/U8.png","sha256":"aaaa","sizeBytes":2048},
                {"path":"Files/datasheet.pdf","sha256":"bbbb","sizeBytes":9}]}}
            """);

        Assert.Equal(2, detail!.Assets.Files.Count);
        Assert.Equal("Images/U8.png", detail.Assets.Files[0].Path);
        Assert.Equal("aaaa", detail.Assets.Files[0].Sha256);
    }

    [Fact]
    public void A_file_missing_its_PATH_or_HASH_is_skipped_rather_than_defaulted()
    {
        // A blank hash builds a fetch URL the server refuses, and the panel would then report a
        // failure that is really a malformed manifest. Skipping keeps the rest reviewable - the
        // same rule the queue parser follows for one bad row.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "manifest":{"files":[
                {"path":"","sha256":"aaaa"},
                {"path":"b.png"},
                {"path":"good.png","sha256":"cccc"}]}}
            """);

        ReviewSubmittedFile file = Assert.Single(detail!.Assets.Files);
        Assert.Equal("good.png", file.Path);
    }

    [Fact]
    public void A_ROWS_ONLY_submission_yields_an_EMPTY_asset_list_rather_than_null()
    {
        // The commonest contribution there is - a typo fix uploads nothing. A null here would
        // make every caller guard against a case that simply means "no files".
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],"manifest":{"files":[]}}
            """);

        Assert.NotNull(detail!.Assets);
        Assert.Empty(detail.Assets.Files);
    }

    [Fact]
    public void The_PUBLISHED_file_list_is_carried_through()
    {
        // *** THIS FIELD EXISTS BECAUSE DERIVING IT CLIENT-SIDE WAS WRONG. *** The first version of
        // the review window built the published side from the change summary's row keys. Those are
        // NATURAL keys - for an image, BoardLabel|Region|Pin|Name - not file paths, so nothing
        // matched: every submitted image read as an addition and a deletion could never appear.
        // The panel drew perfectly and described the wrong thing, which is why the server now
        // states the list outright.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedFiles":["Images/U8.png","Images/gone.png"]}
            """);

        Assert.Equal(["Images/U8.png", "Images/gone.png"], detail!.PublishedFiles);
    }

    [Fact]
    public void A_response_with_no_published_file_list_yields_an_EMPTY_one()
    {
        // A new system has no published side at all, which correctly makes every image an
        // addition rather than throwing.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[]}
            """);

        Assert.Empty(detail!.PublishedFiles);
    }

    [Fact]
    public void A_published_list_and_a_submission_PAIR_UP_into_a_real_comparison()
    {
        // End to end through the two pieces that have to agree: the parser reads both lists off
        // one response, and the planner pairs them. This is what would have caught the natural-key
        // mistake - the pairing only works when both sides speak in FILE PATHS.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedFiles":["Images/U8.png","Images/gone.png"],
             "manifest":{"files":[
                {"path":"Images/U8.png","sha256":"aaaa"},
                {"path":"Images/new.png","sha256":"bbbb"}]}}
            """);

        IReadOnlyList<ReviewImagePair> pairs =
            ReviewImageComparison.Plan(detail!.Assets, detail.PublishedFiles);

        Assert.Equal(3, pairs.Count);

        // The removal leads, and each of the three outcomes is represented exactly once.
        Assert.Equal(ReviewImageChange.Removed, pairs[0].Change);
        Assert.Equal("Images/gone.png", pairs[0].Path);

        Assert.Contains(pairs, pair =>
            pair.Path == "Images/U8.png" && pair.Change == ReviewImageChange.Replaced);

        Assert.Contains(pairs, pair =>
            pair.Path == "Images/new.png" && pair.Change == ReviewImageChange.Added);
    }

    // ###########################################################################################
    // *** THE PUBLISHED HASHES, AND THE SCREEN THEY FIX (2026-09-23). ***
    //
    // A submission names every file the board references, because the manifest describes the
    // whole intended state. Paired against the published list without hashes, every one of those
    // files is "on both sides" and therefore drawn as replaced - so a one-line description edit
    // on the C64 250407 was shown as "1178 images to compare", all identical. The planner already
    // dropped a matching pair when handed the published hash; the server simply never sent one.
    // ###########################################################################################
    [Fact]
    public void The_PUBLISHED_HASHES_are_carried_through()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedHashes":{"Images/U8.png":"aaaa","Images/U9.png":"bbbb"}}
            """);

        Assert.Equal("aaaa", detail!.PublishedHashes["Images/U8.png"]);
        Assert.Equal("bbbb", detail.PublishedHashes["Images/U9.png"]);
    }

    [Fact]
    public void A_response_with_no_published_hashes_yields_an_EMPTY_map_not_a_failure()
    {
        // An older server that does not send the field must degrade to exactly the behaviour it
        // had before - every shared path shown as replaced - rather than refusing the response.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[]}
            """);

        Assert.NotNull(detail);
        Assert.Empty(detail!.PublishedHashes);
    }

    [Fact]
    public void A_malformed_hash_entry_is_skipped_and_the_rest_are_kept()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedHashes":{"Images/U8.png":"aaaa","Images/bad.png":42,"Images/blank.png":""}}
            """);

        Assert.Single(detail!.PublishedHashes);
        Assert.Equal("aaaa", detail.PublishedHashes["Images/U8.png"]);
    }

    [Fact]
    public void Hash_keys_are_matched_ORDINALLY_like_the_file_list_they_pair_against()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedHashes":{"Images/U8.png":"aaaa"}}
            """);

        Assert.False(detail!.PublishedHashes.ContainsKey("images/u8.png"));
    }

    // ###########################################################################################
    // *** THE REPORTED CASE, END TO END: a rows-only change yields NO image pairs at all. ***
    //
    // The three pieces that have to agree - the parser reads the files, the hashes and the
    // submission off one response, and the planner pairs them. This is exactly what the maintainer
    // saw fail: with the hashes absent the same response produces one pair per shared file.
    // ###########################################################################################
    [Fact]
    public void A_submission_whose_files_are_all_byte_identical_produces_NO_pairs()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedFiles":["Images/U8.png","Images/U9.png"],
             "publishedHashes":{"Images/U8.png":"aaaa","Images/U9.png":"bbbb"},
             "manifest":{"files":[
                {"path":"Images/U8.png","sha256":"aaaa"},
                {"path":"Images/U9.png","sha256":"bbbb"}]}}
            """);

        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            detail!.Assets, detail.PublishedFiles, detail.PublishedHashes);

        Assert.Empty(pairs);
    }

    [Fact]
    public void With_hashes_present_only_the_file_that_REALLY_changed_is_paired()
    {
        // The anti-vacuity half: the hashes must not simply switch the comparison off. A file
        // whose submitted hash differs is still shown, as a replacement, and only that one.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "publishedFiles":["Images/U8.png","Images/U9.png"],
             "publishedHashes":{"Images/U8.png":"aaaa","Images/U9.png":"bbbb"},
             "manifest":{"files":[
                {"path":"Images/U8.png","sha256":"aaaa"},
                {"path":"Images/U9.png","sha256":"changed"}]}}
            """);

        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            detail!.Assets, detail.PublishedFiles, detail.PublishedHashes);

        ReviewImagePair pair = Assert.Single(pairs);
        Assert.Equal("Images/U9.png", pair.Path);
        Assert.Equal(ReviewImageChange.Replaced, pair.Change);
    }

    [Fact]
    public void The_SCHEMATIC_IMAGE_map_is_carried_through()
    {
        // A highlight names a SCHEMATIC, not a file - its row key is SchematicName|BoardLabel. So
        // without this map the app knows a rectangle moved on "Sheet 1" and cannot find the
        // picture of Sheet 1 to draw it on.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "schematicImages":{"Sheet 1":"Images/sheet1.png","Sheet 2":"Images/sheet2.png"}}
            """);

        Assert.Equal("Images/sheet1.png", detail!.SchematicImages["Sheet 1"]);
        Assert.Equal(2, detail.SchematicImages.Count);
    }

    [Fact]
    public void A_response_with_no_schematic_images_yields_an_EMPTY_map()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[]}
            """);

        Assert.Empty(detail!.SchematicImages);
    }

    [Fact]
    public void Schematic_names_are_matched_CASE_SENSITIVELY()
    {
        // These keys are matched against a row key the SERVER built, on a tree that is
        // case-sensitive everywhere else. Folding case would pair a highlight with the wrong
        // board.
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"submission":{"id":1},"findings":[],
             "schematicImages":{"Sheet 1":"a.png"}}
            """);

        Assert.True(detail!.SchematicImages.ContainsKey("Sheet 1"));
        Assert.False(detail.SchematicImages.ContainsKey("sheet 1"));
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
        Assert.Empty(detail!.Assets.Files);
    }
}
