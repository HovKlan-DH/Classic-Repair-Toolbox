using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// How "Submit this model" reads a REFUSAL's sentences (SubmissionClient.ReadProblems - the
// oscilloscope model route answers { "errors": [ "..." ] }). What it finds is shown to the user as
// the server's own words; what it cannot read is no sentence, and the caller then says the status
// instead. It never throws: the send it is part of answers every failure in words (code review,
// 2026-10-29 - a body that was JSON but not an object threw out of it).
// ###########################################################################################
public sealed class SubmissionClientRefusalTests
{
    [Fact]
    public void The_server_s_sentences_are_read_in_order()
    {
        Assert.Equal(
            ["Identify is needed.", "That does not look like an email address."],
            SubmissionClient.ReadProblems("{\"errors\":[\"Identify is needed.\",\"That does not look like an email address.\"]}"));
    }

    // Only sentences: a number, an object or an empty string in the list is left out.
    [Fact]
    public void Only_the_strings_of_the_list_with_something_in_them_count()
    {
        Assert.Equal(
            ["Identify is needed."],
            SubmissionClient.ReadProblems("{\"errors\":[42,\"\",{\"x\":1},null,\"Identify is needed.\"]}"));
    }

    [Theory]
    [InlineData("{\"message\":\"Too many requests.\"}")]
    [InlineData("{\"errors\":\"Identify is needed.\"}")]
    [InlineData("{\"errors\":null}")]
    public void A_body_without_a_list_of_errors_has_no_sentence(string body)
    {
        Assert.Empty(SubmissionClient.ReadProblems(body));
    }

    // A proxy's page, a cut-off answer, or JSON that is not an object - none is a sentence, and
    // none throws.
    [Theory]
    [InlineData("")]
    [InlineData("<html><body>502 Bad Gateway</body></html>")]
    [InlineData("{\"errors\":[\"Identify is")]
    [InlineData("[\"Identify is needed.\"]")]
    [InlineData("\"Identify is needed.\"")]
    [InlineData("42")]
    public void A_body_that_is_not_an_object_is_no_sentence_and_never_throws(string body)
    {
        Assert.Empty(SubmissionClient.ReadProblems(body));
    }
}
