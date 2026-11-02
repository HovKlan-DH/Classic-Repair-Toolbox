using System.Globalization;
using System.Net.Http;
using ClassicRepairToolbox.Tests.Maintainer;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// How one file goes up. *** A FILE OF NO BYTES NEVER GOES (owner decision, 2026-11-02: "Refuse empty
// files"): *** hashing says so before anything is sent (SubmissionClientHashFilesTests) and the server
// refuses it at create (SubmissionFlowTests). One emptied between the hashing and its upload is said here,
// for what it is: with no PUT sent, it read "The file [x] has not finished uploading" on every retry.
//
// No network: the answers are built in memory (test rule 6).
// ###########################################################################################
public sealed class SubmissionClientUploadTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private sealed record Part(long Offset, byte[] Body);

    // A server that says `uploaded` bytes have arrived, and takes every part.
    private static (HttpClient Http, List<Part> Parts) Server(long uploaded = 0)
    {
        var parts = new List<Part>();

        var handler = new AnsweringHttpHandler(async (request, token) =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return AnsweringHttpHandler.Json($"{{\"uploaded\":{uploaded.ToString(CultureInfo.InvariantCulture)},\"complete\":false}}");
            }

            string offset = request.RequestUri!.Query.Split('=')[1];
            parts.Add(new Part(long.Parse(offset, CultureInfo.InvariantCulture), await request.Content!.ReadAsByteArrayAsync(token)));

            return AnsweringHttpHandler.Json("{\"uploaded\":0,\"complete\":true}");
        });

        return (new HttpClient(handler), parts);
    }

    private async Task<List<Part>> UploadAsync(byte[] bytes, long uploaded = 0)
    {
        string file = this.thisWorkspace.Path_("file.bin");
        await File.WriteAllBytesAsync(file, bytes);

        (HttpClient http, List<Part> parts) = Server(uploaded);

        using (http)
        {
            await new SubmissionClient("https://server.invalid/api").UploadBlobAsync(http, 7, "token", "abc", file);
        }

        return parts;
    }

    [Fact]
    public async Task A_file_emptied_since_it_was_hashed_is_said_as_empty_and_sends_nothing()
    {
        string file = this.thisWorkspace.Path_("notes.txt");
        await File.WriteAllBytesAsync(file, []);

        (HttpClient http, List<Part> parts) = Server();

        using (http)
        {
            var refused = await Assert.ThrowsAsync<SubmissionRejectedException>(
                () => new SubmissionClient("https://server.invalid/api").UploadBlobAsync(http, 7, "token", "abc", file));

            Assert.Equal(SubmissionFileRules.EmptyFileMessage("notes.txt"), refused.Message);
        }

        Assert.Empty(parts);
    }

    [Fact]
    public async Task A_file_with_bytes_goes_with_them()
    {
        Part part = Assert.Single(await this.UploadAsync([1, 2, 3]));

        Assert.Equal(0, part.Offset);
        Assert.Equal([1, 2, 3], part.Body);
    }

    [Fact]
    public async Task A_file_resumed_sends_only_what_is_left()
    {
        Part part = Assert.Single(await this.UploadAsync([1, 2, 3], uploaded: 2));

        Assert.Equal(2, part.Offset);
        Assert.Equal([3], part.Body);
    }

    // As before: all its bytes arrived, so a resume sends nothing.
    [Fact]
    public async Task A_file_with_every_byte_arrived_sends_nothing()
    {
        Assert.Empty(await this.UploadAsync([1, 2, 3], uploaded: 3));
    }
}
