using System;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// CrtWaitWording - what CRT's "please wait" overlay says, and what the user is told when a wait ran
// into the two-minute limit (owner decision, 2026-09-28). Written for a hobbyist at the bench: each
// sentence names what is happening, and after the limit says what is known and what to do.
// ###########################################################################################
public sealed class CrtWaitWordingTests
{
    // "Working..." was the line that started this - it said nothing about what was being waited for.
    [Fact]
    public void No_wait_sentence_is_a_bare_working()
    {
        string[] sentences =
        [
            CrtWaitWording.SendingFeedback,
            CrtWaitWording.CheckingSubmissions,
            CrtWaitWording.SavingDraft,
            CrtWaitWording.SavingLabels,
            CrtWaitWording.OpeningTable,
            CrtWaitWording.ComparingWithOfficial,
            CrtWaitWording.AddingSchematics,
            CrtWaitWording.ImportingKiCad,
            CrtWaitWording.CheckingKiCadMatches,
            CrtWaitWording.ConnectingOscilloscope,
            CrtWaitWording.DownloadingUpdate,
            CrtWaitWording.ExportingWorkbook(asZip: true),
            CrtWaitWording.ExportingWorkbook(asZip: false),
        ];

        foreach (string sentence in sentences)
        {
            Assert.False(string.IsNullOrWhiteSpace(sentence));
            Assert.DoesNotContain("Working", sentence, StringComparison.Ordinal);
            Assert.EndsWith("...", sentence, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void The_export_names_its_format()
    {
        Assert.Contains("ZIP", CrtWaitWording.ExportingWorkbook(asZip: true), StringComparison.Ordinal);
        Assert.Contains("PDF", CrtWaitWording.ExportingWorkbook(asZip: false), StringComparison.Ordinal);
    }

    [Fact]
    public void Counted_sentences_say_how_far_along_they_are()
    {
        Assert.Equal("Checking your contributions with the server (2 of 5)...", CrtWaitWording.CheckingSubmission(2, 5));
        Assert.Equal("Downloading the update: 40%", CrtWaitWording.DownloadingUpdateAt(40));
        Assert.Equal("Sending to server... 40%", CrtWaitWording.SendingFeedbackAt(40));
    }

    // Once every byte is sent, the server still unpacks and mails before it answers - the line says
    // that is what is waited for, rather than sitting on "100%" (code review, 2026-10-04).
    [Fact]
    public void A_fully_sent_feedback_says_it_waits_for_the_server()
    {
        Assert.Equal("Sent - waiting for the server to save the files and mail the feedback...", CrtWaitWording.SendingFeedbackAt(100));
        Assert.DoesNotContain("100%", CrtWaitWording.SendingFeedbackAt(100), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** FEEDBACK CANNOT BE CHECKED AFTERWARDS, AND SAYS SO. *** The feedback page answers nothing
    // back, so after the limit it MAY have arrived. The sentence must not claim either outcome, must
    // say the text is kept, and must not push a blind resend.
    // ###########################################################################################
    [Fact]
    public void Feedback_with_no_answer_says_it_may_have_arrived_and_the_text_is_kept()
    {
        string text = CrtWaitWording.FeedbackNoAnswer;

        Assert.Contains("within 2 minutes", text, StringComparison.Ordinal);
        Assert.Contains("may still have arrived", text, StringComparison.Ordinal);
        Assert.Contains("kept", text, StringComparison.Ordinal);
        Assert.DoesNotContain("failed", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Each_after_the_limit_sentence_names_the_limit()
    {
        foreach (string text in new[] { CrtWaitWording.SubmissionsNoAnswer, CrtWaitWording.OscilloscopeNoAnswer, CrtWaitWording.UpdateNoProgress })
            Assert.Contains("2 minutes", text, StringComparison.Ordinal);
    }
}
