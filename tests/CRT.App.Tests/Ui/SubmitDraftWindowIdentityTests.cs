using System.Reflection;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Covers SubmitDraftWindow.IdentityToSend - the identity the manifest is built from at the moment
// of sending: the one the window was opened with, plus the summary typed since and the time.
//
// *** EVERY OTHER PROPERTY IS COPIED, CHECKED BY REFLECTION (2026-10-05). *** The copy was written
// out field by field, so a field added to SubmissionIdentity reached the window and stopped there:
// a new system's notes from "Create system" were filled in by TabDrafts and never left the
// computer. Walking the properties makes the next new one fail here instead.
//
// A static method: no window is built, so this is not in the "HeadlessUi" collection.
// ###########################################################################################
public sealed class SubmitDraftWindowIdentityTests
{
    private static readonly DateTimeOffset Opened = new(2026, 10, 1, 9, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset Sending = new(2026, 10, 5, 12, 0, 0, TimeSpan.Zero);

    private static readonly PropertyInfo[] Properties = typeof(SubmissionIdentity)
        .GetProperties(BindingFlags.Public | BindingFlags.Instance);

    // Every string property gets its own name as its value, so a swapped pair fails too.
    private static SubmissionIdentity EveryPropertySet()
    {
        var identity = new SubmissionIdentity();

        foreach (PropertyInfo property in SubmitDraftWindowIdentityTests.Properties)
        {
            if (property.PropertyType == typeof(string))
                property.SetValue(identity, $"value of {property.Name}");
            else if (property.PropertyType == typeof(DateTimeOffset))
                property.SetValue(identity, SubmitDraftWindowIdentityTests.Opened);
            else
                Assert.Fail($"SubmissionIdentity.{property.Name} is a {property.PropertyType.Name}, which this test does not fill in yet.");
        }

        return identity;
    }

    [Fact]
    public void Every_property_but_the_summary_and_the_time_is_copied()
    {
        SubmissionIdentity opened = SubmitDraftWindowIdentityTests.EveryPropertySet();

        SubmissionIdentity sent = SubmitDraftWindow.IdentityToSend(opened, "Typed in the window.", SubmitDraftWindowIdentityTests.Sending);

        foreach (PropertyInfo property in SubmitDraftWindowIdentityTests.Properties)
        {
            object? expected = property.Name switch
            {
                nameof(SubmissionIdentity.Summary) => "Typed in the window.",
                nameof(SubmissionIdentity.CreatedUtc) => SubmitDraftWindowIdentityTests.Sending,
                _ => property.GetValue(opened)
            };

            Assert.True(Equals(expected, property.GetValue(sent)), $"SubmissionIdentity.{property.Name} was not carried into the submission.");
        }
    }

    // The case that found it: a new system's notes reach the identity the manifest is built from.
    [Fact]
    public void A_new_systems_notes_are_sent()
    {
        SubmissionIdentity sent = SubmitDraftWindow.IdentityToSend(
            new SubmissionIdentity { SystemId = "Commodore/C128/310378 Open128", HardwareNotes = "Open-source replica." },
            "A new board.",
            SubmitDraftWindowIdentityTests.Sending);

        Assert.Equal("Open-source replica.", sent.HardwareNotes);
    }
}
