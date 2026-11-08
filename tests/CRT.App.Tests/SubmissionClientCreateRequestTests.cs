using System.Net.Http;
using System.Text.Json;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Covers SubmissionClient.BuildCreateRequest - the request that sends a submission's manifest,
// built apart from sending it so no test touches the network (rule 6).
//
// *** THE AUTHORIZATION HEADER IS A SIGNED-IN MAINTAINER'S SESSION (owner request, 2026-10-01). ***
// With it the server makes the submission the account's, its address verified - the line the
// owner found confusing on their own submission was "Sent without an account". Without a token
// the request must carry NO such header at all: an ordinary contributor has no account, and a
// stray empty "Bearer" would be a header the server has to second-guess.
// ###########################################################################################
public sealed class SubmissionClientCreateRequestTests
{
    private static SubmissionManifest Manifest() => new()
    {
        BoardId = "Commodore/C64/250407",
        Manufacturer = "Commodore",
        Hardware = "C64",
        Board = "250407",
        Summary = "Fixed U8",
        ContactEmail = "dh@example.com"
    };

    [Fact]
    public void A_signed_in_maintainers_token_goes_as_a_bearer_token()
    {
        using HttpRequestMessage request = SubmissionClient.BuildCreateRequest(
            "https://example.com/api/", SubmissionClientCreateRequestTests.Manifest(), " the-token ");

        Assert.Equal(HttpMethod.Post, request.Method);
        Assert.Equal("https://example.com/api/submissions", request.RequestUri!.ToString());
        Assert.Equal("Bearer", request.Headers.Authorization!.Scheme);
        Assert.Equal("the-token", request.Headers.Authorization.Parameter);
    }

    [Fact]
    public void Without_a_token_there_is_no_authorization_header()
    {
        using HttpRequestMessage none = SubmissionClient.BuildCreateRequest(
            "https://example.com/api", SubmissionClientCreateRequestTests.Manifest(), null);

        using HttpRequestMessage blank = SubmissionClient.BuildCreateRequest(
            "https://example.com/api", SubmissionClientCreateRequestTests.Manifest(), "  ");

        Assert.Null(none.Headers.Authorization);
        Assert.Null(blank.Headers.Authorization);
    }

    // The manifest goes as the server reads it - camelCase JSON, as PostAsJsonAsync sent it before.
    [Fact]
    public async Task The_manifest_is_the_body_in_camel_case()
    {
        using HttpRequestMessage request = SubmissionClient.BuildCreateRequest(
            "https://example.com/api", SubmissionClientCreateRequestTests.Manifest(), null);

        using JsonDocument body = JsonDocument.Parse(await request.Content!.ReadAsStringAsync());

        Assert.Equal("Commodore/C64/250407", body.RootElement.GetProperty("boardId").GetString());
        Assert.Equal("dh@example.com", body.RootElement.GetProperty("contactEmail").GetString());
    }

    // ###########################################################################################
    // A new board's notes (owner request, 2026-10-05) go as "hardwareNotes", and come back out of
    // the body as the server binds it - ASP.NET Core's web defaults - so the maintainer's placement
    // can start with them.
    // ###########################################################################################
    [Fact]
    public async Task A_new_boards_notes_go_in_the_body_and_read_back_as_the_server_reads_them()
    {
        SubmissionManifest manifest = SubmissionClientCreateRequestTests.Manifest();
        manifest.HardwareNotes = "Open-source replica.";

        using HttpRequestMessage request = SubmissionClient.BuildCreateRequest("https://example.com/api", manifest, null);
        string json = await request.Content!.ReadAsStringAsync();

        using JsonDocument body = JsonDocument.Parse(json);
        Assert.Equal("Open-source replica.", body.RootElement.GetProperty("hardwareNotes").GetString());

        SubmissionManifest? received = JsonSerializer.Deserialize<SubmissionManifest>(json, new JsonSerializerOptions(JsonSerializerDefaults.Web));
        Assert.Equal("Open-source replica.", received!.HardwareNotes);
    }

    // ###########################################################################################
    // A BOARD FILE'S AND A BOARD LINK'S NOTE (owner request, 2026-11-06) travels the whole way: CRT's
    // body, the server's binding (ASP.NET Core's web defaults), and the publish that writes the rows into
    // the board - so the note typed in the Contribute tab is the note every "Resources" tab shows.
    // ###########################################################################################
    [Fact]
    public async Task A_board_files_and_links_note_reaches_the_published_board()
    {
        SubmissionManifest manifest = SubmissionClientCreateRequestTests.Manifest();
        manifest.Rows.BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf", Note = "Page 12 has the PLA.\nThe rest is the C64C." }];
        manifest.Rows.BoardLinks = [new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://lemon64.com", Note = "Search for 250407" }];

        using HttpRequestMessage request = SubmissionClient.BuildCreateRequest("https://example.com/api", manifest, null);
        string json = await request.Content!.ReadAsStringAsync();

        SubmissionManifest received = JsonSerializer.Deserialize<SubmissionManifest>(json, new JsonSerializerOptions(JsonSerializerDefaults.Web))!;
        BoardData published = PublishMerge.Build(received, published: null);

        Assert.Equal("Page 12 has the PLA.\nThe rest is the C64C.", published.BoardLocalFiles.Single().Note);
        Assert.Equal("Search for 250407", published.BoardLinks.Single().Note);
    }
}
