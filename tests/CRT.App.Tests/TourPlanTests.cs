using Handlers.Tours;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// TourPlan - which steps of a tab's introduction tour to show, given what was seen there (owner request, 2026-11-11:
// "How to handle updates, and then new functionalities? ... what should happen if one update and has skipped multiple
// previous versions?"). Never seen: the introduction. Seen at an older version: what EVERY version since brought, in
// the tour's own order. Nothing new: nothing.
// ###########################################################################################
public sealed class TourPlanTests
{
    private static TourStep Step(string since, string headline, bool inIntroduction = true) =>
        new(since, ["Somewhere"], headline, "Body", inIntroduction);

    // A tab whose tour grew over three releases, one step of 4.1.0 shown to those who update only.
    private static readonly TourOfTab Grown = new(
        TourTab.Schematics,
        "Schematics",
        "Purpose.",
        [
            Step("4.0.0", "Pick a board"),
            Step("4.2.0", "Regions of its own"),
            Step("4.0.0", "Find a component"),
            Step("4.1.0", "A small change", inIntroduction: false),
            Step("4.1.0", "Draw a trace"),
        ]);

    private static string[] Headlines(TourPlan plan) => plan.Steps.Select(step => step.Headline).ToArray();

    [Fact]
    public void Never_seen_is_the_introduction_of_every_version_without_the_steps_for_updates_only()
    {
        TourPlan plan = TourPlan.For(Grown, seen: null);

        Assert.Equal(TourKind.Introduction, plan.Kind);
        Assert.Equal(["Pick a board", "Regions of its own", "Find a component", "Draw a trace"], Headlines(plan));
        Assert.Equal("4.2.0", plan.VersionToRemember);
        Assert.Empty(plan.NewIn);
    }

    // Garbage in the settings file is no version: the introduction, as if nothing were stored.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("seen")]
    [InlineData("4.0")]
    public void A_stored_value_that_is_no_version_is_never_seen(string stored)
    {
        Assert.Equal(TourKind.Introduction, TourPlan.For(Grown, stored).Kind);
    }

    [Fact]
    public void Seen_at_the_newest_version_shows_nothing()
    {
        TourPlan plan = TourPlan.For(Grown, "4.2.0");

        Assert.Equal(TourKind.None, plan.Kind);
        Assert.Empty(plan.Steps);
        Assert.Equal("4.2.0", plan.VersionToRemember);
    }

    [Fact]
    public void Seen_one_version_back_shows_what_the_newest_brought()
    {
        TourPlan plan = TourPlan.For(Grown, "4.1.0");

        Assert.Equal(TourKind.WhatIsNew, plan.Kind);
        Assert.Equal(["Regions of its own"], Headlines(plan));
        Assert.Equal(["4.2.0"], plan.NewIn);
        Assert.Equal("4.2.0", plan.VersionToRemember);
    }

    // *** THE OWNER'S QUESTION: versions skipped. *** From 4.0.0 straight to 4.2.0: what 4.1.0 AND 4.2.0 brought, in the
    // tour's own order - not grouped by version - and a step for updates only is in it.
    [Fact]
    public void Versions_skipped_show_what_every_one_of_them_brought_in_the_tours_order()
    {
        TourPlan plan = TourPlan.For(Grown, "4.0.0");

        Assert.Equal(TourKind.WhatIsNew, plan.Kind);
        Assert.Equal(["Regions of its own", "A small change", "Draw a trace"], Headlines(plan));
        Assert.Equal(["4.1.0", "4.2.0"], plan.NewIn);
        Assert.Equal("4.2.0", plan.VersionToRemember);
    }

    // A pre-release that saw the tour counts as before its release: a step of 4.0.0 is new to somebody who last saw
    // 4.0.0-beta.2 - while the steps keep the release's number, as the catalogue's do, it never is to 4.0.0 itself.
    [Fact]
    public void A_pre_release_ranks_before_its_release()
    {
        Assert.Equal(TourKind.WhatIsNew, TourPlan.For(Grown, "4.2.0-beta.2").Kind);
        Assert.Equal(["Regions of its own"], Headlines(TourPlan.For(Grown, "4.2.0-beta.2")));
    }

    // A CRT run after a newer one (somebody stepped back a release) shows nothing, and never lowers what is remembered.
    [Fact]
    public void Seen_by_a_newer_CRT_shows_nothing_and_keeps_the_newer_version()
    {
        TourPlan plan = TourPlan.For(Grown, "5.0.0");

        Assert.Equal(TourKind.None, plan.Kind);
        Assert.Equal("5.0.0", plan.VersionToRemember);
    }

    // A step whose version is a typo would be "new" at every launch if it were counted - it is never new, and the
    // introduction still has it.
    [Fact]
    public void A_step_with_no_version_is_never_new()
    {
        var tour = new TourOfTab(TourTab.Contribute, "Contribute", "Purpose.", [Step("4.0.0", "One"), Step("four", "Typo")]);

        Assert.Equal(TourKind.None, TourPlan.For(tour, "4.0.0").Kind);
        Assert.Equal(["One", "Typo"], Headlines(TourPlan.For(tour, null)));
        Assert.Equal("4.0.0", TourPlan.For(tour, null).VersionToRemember);
    }

    [Fact]
    public void A_tour_with_only_steps_for_updates_has_no_introduction()
    {
        var tour = new TourOfTab(TourTab.Workbooks, "Workbooks", "Purpose.", [Step("4.1.0", "Later", inIntroduction: false)]);

        Assert.Equal(TourKind.None, TourPlan.For(tour, null).Kind);
        Assert.Equal(TourKind.WhatIsNew, TourPlan.For(tour, "4.0.0").Kind);
    }

    // ###########################################################################################
    // *** WHAT A SHOWING LEFT OUT COMES BACK (code review, 2026-11-12: "Export to PDF" never reached a new user with no
    // workbook, the trace step one whose first schematic image had no KiCad data). *** Seen at the newest version, a
    // step left out is shown alone, as LeftOut; with steps that are new, among them in the tour's order - never badged
    // as new. A headline no step has any more is nothing.
    // ###########################################################################################
    [Fact]
    public void A_step_left_out_comes_back_alone_or_among_the_new_ones()
    {
        TourPlan alone = TourPlan.For(Grown, "4.2.0", ["Find a component", "Gone from the catalogue"]);

        Assert.Equal(TourKind.LeftOut, alone.Kind);
        Assert.Equal(["Find a component"], Headlines(alone));
        Assert.Equal(alone.Steps, alone.LeftOutBefore);
        Assert.False(alone.IsNew(alone.Steps[0]));
        Assert.Equal("4.2.0", alone.VersionToRemember);

        TourPlan among = TourPlan.For(Grown, "4.1.0", ["Find a component"]);

        Assert.Equal(TourKind.WhatIsNew, among.Kind);
        Assert.Equal(["Regions of its own", "Find a component"], Headlines(among));
        Assert.True(among.IsNew(among.Steps[0]));
        Assert.False(among.IsNew(among.Steps[1]));
        Assert.Equal(["4.2.0"], among.NewIn);
    }

    // The introduction has every step anyway: what an earlier showing left out changes nothing.
    [Fact]
    public void An_introduction_ignores_what_was_left_out()
    {
        Assert.Equal(Headlines(TourPlan.For(Grown, null)), Headlines(TourPlan.For(Grown, null, ["Draw a trace"])));
        Assert.Empty(TourPlan.For(Grown, null, ["Draw a trace"]).LeftOutBefore);
    }

    // What to remember as the showing ends: every step of the plan not said whole - the ones off screen and the ones
    // said in part alike - and nothing it said whole, a step left out before included.
    [Fact]
    public void What_is_left_out_after_a_showing_is_every_step_not_said_whole()
    {
        TourPlan introduction = TourPlan.For(Grown, null);
        IReadOnlyList<TourStep> steps = introduction.Steps;

        Assert.Equal(["Regions of its own", "Draw a trace"], introduction.LeftOutAfter([steps[0], steps[2]]));
        Assert.Empty(introduction.LeftOutAfter(steps));

        TourPlan alone = TourPlan.For(Grown, "4.2.0", ["Find a component"]);
        Assert.Empty(alone.LeftOutAfter(alone.Steps));
        Assert.Equal(["Find a component"], alone.LeftOutAfter([]));
    }

    // The first card names only the versions of steps really shown - one off screen takes its version with it.
    [Fact]
    public void Only_the_versions_of_the_steps_shown_are_named()
    {
        TourPlan plan = TourPlan.For(Grown, "4.0.0");

        Assert.Equal(["4.1.0"], plan.NewInOf(plan.Steps.Where(step => step.Since == "4.1.0")));
        Assert.Equal(["4.1.0", "4.2.0"], plan.NewInOf(plan.Steps));
        Assert.Empty(plan.NewInOf([]));
    }
}
