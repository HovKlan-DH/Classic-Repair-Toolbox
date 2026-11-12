using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Tours
{
    internal enum TourKind
    {
        // Nothing to show: seen, and nothing new since.
        None,

        // Never seen on this machine: the tab's introduction.
        Introduction,

        // Seen at an older version: what came since, of every version in between.
        WhatIsNew,

        // Seen, nothing new - but a step an earlier showing left out, or said only in part, as what it is about was not
        // on screen then (code review, 2026-11-12).
        LeftOut,
    }

    // ###########################################################################################
    // Which steps of a tab's tour to show, given the version of it last seen there (TourCatalogue's header says how a
    // tour grows with CRT) and the steps an earlier showing left out. Pure: what is stored is handed in.
    //
    //   - nothing stored, or what is stored no version at all: the introduction, every InIntroduction step;
    //   - a version: every step whose Since is newer, in the tour's own order - "Tour of what's new";
    //   - with it, every step an earlier showing LEFT OUT, by its headline (UserSettings.IntroductionTourStepsLeftOut):
    //     a step is shown only while what it frames is on screen, and a paragraph only while its control is
    //     (Main.Tours.cs) - so a new user with no workbook was never told of "Export to PDF", nor one whose first
    //     schematic image had no KiCad data of highlighting a trace (code review, 2026-11-12). Such a step comes back
    //     once it can be said WHOLE (LeftOutBefore), in the tour's order among the new ones, or alone as LeftOut;
    //   - VersionToRemember is the newest Since of the tour, never older than what was stored, so a CRT run after a newer
    //     one (a step back to an older release) neither lowers it nor shows the older tour again.
    //
    // A step whose Since is no version is never "new": it would otherwise show at every launch.
    // ###########################################################################################
    internal sealed record TourPlan(TourKind Kind, IReadOnlyList<TourStep> Steps, string VersionToRemember, IReadOnlyList<string> NewIn)
    {
        // The steps of Steps an earlier showing left out - shown only when they can be said whole.
        public IReadOnlyList<TourStep> LeftOutBefore { get; init; } = [];

        // Whether `step` is in this plan for being new - its card names the version it came with.
        public bool IsNew(TourStep step) => this.Kind == TourKind.WhatIsNew && this.Steps.Contains(step) && !this.LeftOutBefore.Contains(step);

        // ###########################################################################################
        // What to remember as left out once a showing of this plan ends: every step of it NOT said whole - `shownWhole`
        // are those that were. By headline, which TourCatalogueTests holds unique per tour.
        // ###########################################################################################
        public IReadOnlyList<string> LeftOutAfter(IEnumerable<TourStep> shownWhole)
        {
            HashSet<TourStep> said = shownWhole.ToHashSet();
            return this.Steps.Where(step => !said.Contains(step)).Select(step => step.Headline).ToList();
        }

        public static TourPlan For(TourOfTab tour, string? seen, IReadOnlyCollection<string>? leftOut = null)
        {
            CrtVersion? newest = tour.Steps
                .Select(step => CrtVersion.TryParse(step.Since, out CrtVersion? version) ? version : null)
                .Where(version => version is not null)
                .Max();

            if (!CrtVersion.TryParse(seen, out CrtVersion? seenVersion) || seenVersion is null)
            {
                List<TourStep> introduction = tour.Steps.Where(step => step.InIntroduction).ToList();

                return new TourPlan(
                    introduction.Count == 0 ? TourKind.None : TourKind.Introduction,
                    introduction,
                    newest?.ToString() ?? string.Empty,
                    []);
            }

            List<TourStep> newer = tour.Steps
                .Where(step => CrtVersion.TryParse(step.Since, out CrtVersion? since) && since! > seenVersion)
                .ToList();

            List<TourStep> before = tour.Steps
                .Where(step => !newer.Contains(step) && leftOut is not null && leftOut.Contains(step.Headline, StringComparer.Ordinal))
                .ToList();

            IReadOnlyList<string> newIn = newer
                .Select(step => CrtVersion.Parse(step.Since))
                .Distinct()
                .Order()
                .Select(version => version.ToString())
                .ToList();

            string remember = newest is not null && newest > seenVersion ? newest.ToString() : seenVersion.ToString();

            TourKind kind = newer.Count > 0 ? TourKind.WhatIsNew : before.Count > 0 ? TourKind.LeftOut : TourKind.None;

            return new TourPlan(kind, tour.Steps.Where(step => newer.Contains(step) || before.Contains(step)).ToList(), remember, newIn)
            {
                LeftOutBefore = before,
            };
        }

        // ###########################################################################################
        // The versions of NewIn that `shown` - the steps really shown, some may be off screen - came with. The first card
        // names no version none of its steps is about.
        // ###########################################################################################
        public IReadOnlyList<string> NewInOf(IEnumerable<TourStep> shown)
        {
            HashSet<CrtVersion> versions = shown
                .Select(step => CrtVersion.TryParse(step.Since, out CrtVersion? since) ? since : null)
                .OfType<CrtVersion>()
                .ToHashSet();

            return this.NewIn.Where(version => versions.Contains(CrtVersion.Parse(version))).ToList();
        }
    }
}
