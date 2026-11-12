using Handlers.DataHandling;
using System.Collections.Generic;

namespace Handlers.Tours
{
    // ###########################################################################################
    // What the tours' cards say around the steps themselves (the steps' own words are TourCatalogue's). The first card
    // says plainly that it IS a tour (owner request, 2026-11-11: "it should clearly show this is an "Introduction
    // tour""), and a tour of what changed names the versions it covers, every one skipped included.
    // ###########################################################################################
    internal static class TourWording
    {
        public const string SkipButton = "Skip tour";
        public const string StartButton = "Start tour";
        public const string BackButton = "Back";
        public const string NextButton = "Next";
        public const string DoneButton = "Done";

        // Where the tours are taken again - the "Configuration" tab's button.
        public const string TakeAgainButton = "Take the introduction tours again";

        // A tab, a button or a panel a tour names is BOLD, never in quotes (owner request, 2026-11-11: "in all
        // introductions, then try and use bold instead of encapsulating in """) - the cards show "<b>...</b>" bold.
        public const string TakeAgainFoot = "You can take the tours again from the <b>Configuration</b> tab.";

        // What an earlier tour left out is the introduction going on, and says so - "Introduction tour" alone, over words
        // saying it shows what an earlier one left out, read as the first look again (code review, 2026-11-12).
        public static string Kicker(TourKind kind) => kind switch
        {
            TourKind.WhatIsNew => "Tour of what's new",
            TourKind.LeftOut => "Introduction tour, continued",
            _ => "Introduction tour",
        };

        // The headline is bold as a whole, so the tab's name stands in it as it is.
        public static string WelcomeHeadline(TourOfTab tour) => $"The {tour.TabName} tab";

        // ###########################################################################################
        // The first card's words: what the tab is for, then what the tour shows - or, of a tour of what changed, that
        // the tab changed and in which versions. No count of the steps (owner request, 2026-11-11: "just show that it
        // shows the most important ones and skip the counter") - the dots on every step say how far it goes.
        //
        // `steps` counts the NEW steps shown, `leftOut` the steps an earlier tour left out, shown now (TourPlan.LeftOut;
        // code review, 2026-11-12).
        // ###########################################################################################
        public static string WelcomeBody(TourOfTab tour, TourKind kind, int steps, IReadOnlyList<string> newIn, int leftOut = 0)
        {
            if (kind == TourKind.WhatIsNew)
            {
                string what = steps == 1 ? "the most important change" : "the most important changes";
                string versions = newIn.Count == 0 ? string.Empty : $", new in CRT {TourWording.Listed(newIn)}";
                string also = leftOut == 0 ? string.Empty : $" {TourWording.AlsoLeftOut}";

                return $"The <b>{tour.TabName}</b> tab has changed since you last used it. This short tour shows {what}{versions}.{also}";
            }

            if (kind == TourKind.LeftOut)
            {
                return $"{tour.Purpose}\n\nThis short tour shows what an earlier tour of this tab left out, as it was not on screen then.";
            }

            string shows = steps == 1
                ? "This short tour shows the most important thing here."
                : "This short tour shows the most important things here. It takes about a minute.";

            return $"{tour.Purpose}\n\n{shows}";
        }

        // ###########################################################################################
        // THE LAST CARD OF THE FIRST INTRODUCTION TAKEN (owner request, 2026-11-11: "As the last screen it should show info
        // that the user now can go to any other tab to see the introduction for that tab"): framing the other tabs, said
        // only while no other tab's tour has been seen - after that, it is known.
        // ###########################################################################################
        public const string OtherToursHeadline = "More tabs have an introduction";

        // The frame is one rectangle over the tab strip, so "Resources" and "Maintainer" stand in it too - the words name none.
        public const string OtherToursBody = "You can now go to the other tabs - most show an introduction of their own the first time you open them.";

        public static string StepCounter(int number, int of) => $"Step {number} of {of}";

        // Over a step of a tour of what changed: the version it came with.
        public static string NewIn(string version) => $"New in {version}";

        // A tour of what changed that also shows what an earlier tour left out says so.
        public const string AlsoLeftOut = "It also shows what an earlier tour of this tab left out, as it was not on screen then.";

        // "4.1.0", "4.1.0 and 4.2.0", "4.1.0, 4.2.0 and 4.3.0" - CRT.Data's one way of listing in a sentence (code review,
        // 2026-11-12: a copy of its own here).
        public static string Listed(IReadOnlyList<string> items) => BoardDataChecks.ListOf(items);
    }
}
