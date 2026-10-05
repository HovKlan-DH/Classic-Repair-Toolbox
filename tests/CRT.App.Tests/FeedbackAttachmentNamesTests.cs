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
