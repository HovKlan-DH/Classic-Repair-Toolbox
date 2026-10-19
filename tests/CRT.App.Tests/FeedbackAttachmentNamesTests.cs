using System.IO.Compression;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The Feedback tab against the server that receives it (2026-10-03, when feedback moved to
// CRT.Server). The tab's two check boxes attach CRT's own files under AppConfig's
// names; the server SHOWS a file in the mail only when its name is in FeedbackContract.InlineFiles
// - so a renamed log file would quietly turn from the mail's first section into an attachment on
// the share. And the address the tab posts to is the route the server maps (RequestBodyLimitsTests
// pins the same path on the server's real route table).
// ###########################################################################################
public sealed class FeedbackAttachmentNamesTests
{
    [Theory]
    [InlineData(AppConfig.LogFileName)]
    [InlineData(AppConfig.CrashFileName)]
    [InlineData(AppConfig.SettingsFileName)]
    [InlineData(AppConfig.TracesFileName)]
    public void Every_file_the_check_boxes_attach_is_one_the_mail_shows(string fileName)
    {
        Assert.Contains(FeedbackContract.InlineFiles, file => file.FileName == fileName);
    }

    [Fact]
    public void Feedback_goes_to_the_servers_feedback_route()
    {
        Assert.Equal("https://classic-repair-toolbox.dk/api/feedback", TabFeedback.FeedbackUrl);
    }

    // ------------------------------------------------------------------ code review, 2026-10-16

    // A picture that cannot be had as the feedback is sent stops the send, named: it was left out,
    // CRT said "sent" and cleared the field, and a picture the user could see never went. A dropped
    // picture goes by its FILE, streamed into the zip - not read into memory here.
    [Fact]
    public async Task A_picture_that_cannot_be_had_is_named_and_a_dropped_one_goes_by_its_file()
    {
        using var workspace = new TempWorkspace();
        string kept = Path.Combine(workspace.Root, "kept.png");
        File.WriteAllBytes(kept, [1, 2, 3]);

        FeedbackPicture pasted = new("pasted-image-1.png", null, new Avalonia.PixelSize(10, 10), () => Task.FromResult(new byte[] { 1, 2 }));
        FeedbackPicture dropped = new("kept.png", null, new Avalonia.PixelSize(10, 10), () => throw new InvalidOperationException("not read for sending"), filePath: kept);
        FeedbackPicture gone = new("shot.png", null, new Avalonia.PixelSize(10, 10), () => Task.FromResult(new byte[] { 9 }), filePath: Path.Combine(workspace.Root, "shot.png"));
        FeedbackPicture failed = new("pasted-image-2.png", null, new Avalonia.PixelSize(10, 10), () => Task.FromException<byte[]>(new InvalidOperationException("encoding")));

        var (ready, unreadable) = await TabFeedback.PreparePicturesAsync([pasted, dropped, gone, failed]);

        Assert.Equal(["pasted-image-1.png", "kept.png"], ready.Select(picture => picture.Name));
        Assert.Equal(new byte[] { 1, 2 }, ready[0].Bytes);
        Assert.Equal((null, kept), (ready[1].Bytes, ready[1].Path));
        Assert.Equal([gone, failed], unreadable);

        Assert.Equal(
            "Not sent: the picture [shot.png] could not be read - its file may have been moved or deleted; " +
            "the pasted picture [pasted-image-2.png] could not be made ready to send. " +
            "Remove them from the text, or put them in again, and send again.",
            TabFeedback.NotReadWording(unreadable));
    }

    // A pasted picture has no file to blame: its words are its own.
    [Fact]
    public void A_picture_that_cannot_be_had_is_named_with_what_to_do()
    {
        Assert.Equal(
            "Not sent: the picture [shot.png] could not be read - its file may have been moved or deleted. Remove it from the text, or put it in again, and send again.",
            FeedbackWording.PicturesNotRead(["shot.png"], []));
        Assert.Equal(
            "Not sent: the pasted picture [pasted-image-1.png] could not be made ready to send. Remove it from the text, or put it in again, and send again.",
            FeedbackWording.PicturesNotRead([], ["pasted-image-1.png"]));
        Assert.Equal(
            "Not sent: the pictures [a.png], [b.png] could not be read - their files may have been moved or deleted. Remove them from the text, or put them in again, and send again.",
            FeedbackWording.PicturesNotRead(["a.png", "b.png"], []));
        Assert.StartsWith("Not sent: the pasted pictures [pasted-image-1.png], [pasted-image-2.png] could not", FeedbackWording.PicturesNotRead([], ["pasted-image-1.png", "pasted-image-2.png"]), StringComparison.Ordinal);
    }

    [Fact]
    public void A_refusal_the_user_can_act_on_says_what_to_do()
    {
        Assert.Contains("250 MB", FeedbackWording.Failed(413), StringComparison.Ordinal);
        Assert.Equal(FeedbackWording.TooLarge, FeedbackWording.Failed(413));
        Assert.Contains("try again in an hour", FeedbackWording.Failed(429), StringComparison.Ordinal);
        Assert.Contains("try again later", FeedbackWording.Failed(502), StringComparison.Ordinal);
        Assert.Contains("no room", FeedbackWording.Failed(507), StringComparison.Ordinal);
        Assert.Contains("HTTP 404", FeedbackWording.Failed(404), StringComparison.Ordinal);
        Assert.Contains("(HTTP 500) - please check the logfile", FeedbackWording.Failed(500), StringComparison.Ordinal);
    }
}
