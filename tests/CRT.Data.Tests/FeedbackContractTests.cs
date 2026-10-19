using System.IO;
using System.Linq;
using System.Net.Http;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The feedback form both ends share (2026-10-03). The server's side - reading this very form - is
// CRT.Server.Tests' FeedbackFormReaderTests; this pins the form's shape and the success rule.
// ###########################################################################################
public sealed class FeedbackContractTests
{
    [Fact]
    public void The_form_carries_the_three_fields_and_the_zip_only_when_there_is_one()
    {
        using MultipartFormDataContent bare = FeedbackContract.BuildForm("a@example.com", "Text", "2.6.0", attachmentZip: null);
        using MultipartFormDataContent empty = FeedbackContract.BuildForm("a@example.com", "Text", "2.6.0", new MemoryStream());
        using MultipartFormDataContent withZip = FeedbackContract.BuildForm("a@example.com", "Text", "2.6.0", new MemoryStream([1, 2, 3]));

        Assert.Equal(["email", "feedback", "version"], bare.Select(part => part.Headers.ContentDisposition!.Name!.Trim('"')));
        Assert.Equal(3, empty.Count());

        HttpContent zip = withZip.Last();
        Assert.Equal(FeedbackContract.AttachmentField, zip.Headers.ContentDisposition!.Name!.Trim('"'));
        Assert.Equal(FeedbackContract.AttachmentFileName, zip.Headers.ContentDisposition.FileName!.Trim('"'));
        Assert.Equal("application/zip", zip.Headers.ContentType!.MediaType);
    }

    // The pictures' widths (2026-10-15) are a part of their own only when there are any - a
    // feedback without pictures is the form every CRT has always sent, part for part.
    [Fact]
    public void The_pictures_widths_are_a_part_only_when_there_are_any()
    {
        using MultipartFormDataContent none = FeedbackContract.BuildForm("a@example.com", "Text", "2.6.0", attachmentZip: null, pictureWidths: "  ");
        using MultipartFormDataContent with = FeedbackContract.BuildForm("a@example.com", "Text", "2.6.0", attachmentZip: null, pictureWidths: "{\"a.png\":640}");

        Assert.Equal(["email", "feedback", "version"], none.Select(part => part.Headers.ContentDisposition!.Name!.Trim('"')));
        Assert.Equal(["email", "feedback", "version", "pictureWidths"], with.Select(part => part.Headers.ContentDisposition!.Name!.Trim('"')));
    }

    // The old feedback address could answer 200 with a warning in the body, so the body is read as
    // well.
    [Theory]
    [InlineData(200, "Success", true)]
    [InlineData(200, "  success\n", true)]
    [InlineData(200, "Warning: something in index.php", false)]
    [InlineData(200, "", false)]
    [InlineData(502, "Success", false)]
    [InlineData(429, null, false)]
    public void Only_a_2xx_saying_Success_is_success(int status, string? body, bool expected)
    {
        Assert.Equal(expected, FeedbackContract.IsSuccess(status, body));
    }

    [Fact]
    public void The_request_limit_leaves_room_for_the_text_beside_the_largest_zip()
    {
        Assert.True(FeedbackContract.MaximumRequestBytes - FeedbackContract.MaximumAttachmentBytes > FeedbackContract.MaximumFeedbackCharacters * 4L);
    }
}
