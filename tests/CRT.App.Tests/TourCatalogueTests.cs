using Handlers.DataHandling;
using Handlers.Theming;
using Handlers.Tours;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// TourCatalogue - every step of every introduction tour (owner request, 2026-11-11: "It should not be everything it
// should introduce, but just the most important topics"). That each control a step names is in CRT's window is
// MainToursTests' (CRT.App.UiTests) - it needs the window.
// ###########################################################################################
public sealed class TourCatalogueTests
{
    // The owner's three tabs first (owner decision, 2026-11-11: "let us start with this"), then every other (the same day:
    // "OK, now do the rest") - one tour each. That every tab of the window has one is MainToursTests'.
    [Fact]
    public void Every_tab_has_one_tour()
    {
        Assert.Equal([TourTab.Schematics, TourTab.Contribute, TourTab.Workbooks], TourCatalogue.All.Take(3).Select(tour => tour.Tab));
        Assert.Equal(Enum.GetValues<TourTab>().Order(), TourCatalogue.All.Select(tour => tour.Tab).Order());

        foreach (TourTab tab in Enum.GetValues<TourTab>())
        {
            Assert.Same(TourCatalogue.All.Single(tour => tour.Tab == tab), TourCatalogue.Of(tab));
        }
    }

    // The most important things, not everything: a tour of more than seven steps is a manual again. Seven since 2026-11-12
    // (owner requests: the "Schematics" tour's refresh icon and trace steps - the trace step shown only with KiCad data).
    // The steps of the tab's own: the first introduction taken ends on one card more, on the other tabs (owner request,
    // 2026-11-11), which is about no step of the tab and is not counted here (code review, 2026-11-12).
    [Fact]
    public void Each_introduction_is_short()
    {
        foreach (TourOfTab tour in TourCatalogue.All)
        {
            int steps = tour.Steps.Count(step => step.InIntroduction);
            Assert.InRange(steps, 1, 7);
        }
    }

    // A step whose version is no version would never count as new - so its change would never be shown to anybody who
    // updates (TourPlan).
    [Fact]
    public void Every_step_names_the_version_it_came_with()
    {
        foreach (TourStep step in TourCatalogue.All.SelectMany(tour => tour.Steps))
        {
            Assert.True(CrtVersion.TryParse(step.Since, out CrtVersion? since) && since is not null && !since.IsPreRelease, $"[{step.Headline}]: [{step.Since}]");
        }
    }

    [Fact]
    public void Every_step_says_something_and_names_its_controls_as_paths()
    {
        foreach (TourOfTab tour in TourCatalogue.All)
        {
            Assert.False(string.IsNullOrWhiteSpace(tour.Purpose));
            Assert.False(string.IsNullOrWhiteSpace(tour.TabName));

            foreach (TourStep step in tour.Steps)
            {
                Assert.False(string.IsNullOrWhiteSpace(step.Headline));
                Assert.False(string.IsNullOrWhiteSpace(step.Body.Whole));

                // Something is said whatever is on screen - a paragraph said always.
                Assert.Contains(step.Body.Parts, part => part.Shows is null);

                foreach (string path in step.Targets)
                {
                    Assert.All(path.Split('/'), name => Assert.Matches("^[A-Za-z][A-Za-z0-9]*$", name));
                }
            }
        }
    }

    // A step left out of a showing is remembered by its headline (TourPlan.LeftOutAfter), so no two steps of a tour share
    // one - else the one would come back for the other.
    [Fact]
    public void No_two_steps_of_a_tour_share_a_headline()
    {
        foreach (TourOfTab tour in TourCatalogue.All)
        {
            Assert.Equal(tour.Steps.Count, tour.Steps.Select(step => step.Headline).Distinct(StringComparer.Ordinal).Count());
        }
    }

    // A paragraph said only while its control is on screen names one its step frames - else the words and the frame
    // could disagree about what is there.
    [Fact]
    public void A_paragraph_said_only_with_its_control_names_one_its_step_frames()
    {
        foreach (TourStep step in TourCatalogue.All.SelectMany(tour => tour.Steps))
        {
            foreach (TourPart part in step.Body.Parts.Where(part => part.Shows is not null))
            {
                Assert.Contains(part.Shows, step.Targets);
            }
        }
    }

    // *** WHAT IS NOT ON SCREEN IS NOT SAID (owner question, 2026-11-11: "Will this step always show the current board
    // selected, as that may not include the regions buttons? Then what?"). *** A board with no regions is told nothing of
    // "PAL" and "NTSC"; what is said always stays, in its order.
    [Fact]
    public void A_paragraph_about_a_control_not_on_screen_is_left_out()
    {
        TourStep buttons = TourCatalogue.Of(TourTab.Schematics).Steps.Single(step => step.Targets.Contains("RegionButtonsGrid"));

        string withRegions = buttons.Body.TextFor(_ => true);
        string withoutRegions = buttons.Body.TextFor(path => path != "RegionButtonsGrid");

        Assert.Contains("<b>PAL</b>", withRegions, StringComparison.Ordinal);
        Assert.DoesNotContain("PAL", withoutRegions, StringComparison.Ordinal);
        Assert.StartsWith("<b>Clear</b>", withoutRegions, StringComparison.Ordinal);
        Assert.EndsWith("easy to spot.", withoutRegions, StringComparison.Ordinal);
        Assert.DoesNotContain("\n\n\n", withoutRegions, StringComparison.Ordinal);
    }

    // Every text a tour can say - the purpose, the headlines, every paragraph.
    private static IEnumerable<string> TextsOf(TourOfTab tour) =>
        new[] { tour.Purpose, tour.TabName }.Concat(tour.Steps.SelectMany(step => new[] { step.Headline, step.Body.Whole }));

    // Plain ASCII punctuation, as every text CRT shows (CLAUDE.md) - a dash or a quote pasted from elsewhere is caught here.
    [Fact]
    public void The_words_are_plain_ASCII()
    {
        Assert.All(TourCatalogue.All.SelectMany(TextsOf), text => Assert.True(text.All(c => c is '\n' || c is >= ' ' and <= '~'), text));
    }

    // *** A TOUR NAMES A TAB, A BUTTON, A PANEL OR A KEY IN BOLD, NEVER IN QUOTES (owner request, 2026-11-11: "in all
    // introductions, then try and use bold instead of encapsulating in """). *** And "highlight", never "light up" or
    // "lit" (owner request, the same day: "Please always use the term "highlight" for it").
    [Fact]
    public void A_tour_names_things_in_bold_and_says_highlight()
    {
        foreach (string text in TourCatalogue.All.SelectMany(TextsOf))
        {
            Assert.DoesNotContain('"', text);
            Assert.DoesNotMatch(@"(?i)\blight(s|ing)?\b[^.]*\bup\b|\blit\b", text);
        }
    }

    // A body says its bold pieces as "<b>...</b>" (BoldRuns, the card's TourOverlay) - each tag with its partner, or a
    // slip would show as plain words that were meant bold. A headline and a tab's purpose are shown as plain text, so a
    // tag there would be read on the card.
    [Fact]
    public void Only_a_body_marks_bold_pieces_and_every_tag_has_its_partner()
    {
        foreach (TourOfTab tour in TourCatalogue.All)
        {
            Assert.DoesNotContain('<', tour.Purpose);

            foreach (TourStep step in tour.Steps)
            {
                Assert.DoesNotContain('<', step.Headline);

                foreach (TourPart part in step.Body.Parts)
                {
                    Assert.Equal(Count(part.Text, "<b>"), Count(part.Text, "</b>"));
                    Assert.DoesNotContain('<', BoldRuns.Plain(part.Text));
                }
            }
        }

        static int Count(string text, string tag) => (text.Length - text.Replace(tag, string.Empty, StringComparison.Ordinal).Length) / tag.Length;
    }

    // *** THE BOARD'S PICTURES ARE "SCHEMATIC IMAGES" (owner request, 2026-11-11: "Always refer to the 'picture' in the
    // 'Schematics' tab as a 'schematic image'"). *** A component's pictures stay pictures ("pinout, pictures and
    // oscilloscope readings" - what its popup shows), so the test is of "the picture", "a picture" and the like.
    [Fact]
    public void The_Schematics_tour_calls_the_board_s_pictures_schematic_images()
    {
        TourOfTab tour = TourCatalogue.Of(TourTab.Schematics);

        foreach (string text in TextsOf(tour))
        {
            Assert.DoesNotMatch(@"(?i)\b(the|a|this|each|every|between|on|of) pictures?\b", text);
        }

        Assert.Contains(tour.Steps, step => step.Body.Whole.Contains("schematic image", StringComparison.Ordinal));
    }
}
