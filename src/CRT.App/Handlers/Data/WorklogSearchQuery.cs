using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One term of a parsed search query, plus where it was found in a matched string.
    // Start/Length index into the ORIGINAL text (not a lowered copy), so a caller can highlight the
    // exact run the user's term matched without re-searching.
    // ###########################################################################################
    public readonly struct WorklogSearchHit
    {
        public WorklogSearchHit(int start, int length)
        {
            this.Start = start;
            this.Length = length;
        }

        public int Start { get; }
        public int Length { get; }
        public int End => this.Start + this.Length;
    }

    // ###########################################################################################
    // One term from a query string: the text to look for, and whether finding it EXCLUDES the
    // record ("-cpu") rather than being required.
    //
    // Quoted terms and bare terms are the same thing once parsed - the quotes only decide where the
    // term ENDS (a quoted term may contain spaces). Matching is always "is this substring present",
    // case-insensitively, so a term matches mid-word: the spec's "p c u" example finds "CPU"
    // because each single character is present, and the quoted "full text" finds "Afull textB"
    // because the run appears contiguously there (it would not match "full-text").
    // ###########################################################################################
    public sealed class WorklogSearchTerm
    {
        public WorklogSearchTerm(string text, bool isExcluded)
        {
            this.Text = text ?? string.Empty;
            this.IsExcluded = isExcluded;
        }

        public string Text { get; }
        public bool IsExcluded { get; }
    }

    // ###########################################################################################
    // A parsed free-text search query for the Workbooks tab's "Find a previous repair" box.
    //
    // Grammar, taken from the feature request verbatim:
    //   "full text"   - a quoted run is ONE term, spaces included. Still a substring match, so it
    //                   also finds "Afull textB".
    //   cpu super     - a space is a logical AND: every term must be found somewhere in the record.
    //   p c u         - falls out of the same rule; single characters are just short terms, so this
    //                   matches "CPU" (each of p, c and u is present).
    //   -cpu          - a leading minus EXCLUDES: the record must NOT contain it. Works on quoted
    //                   terms too ( -"full text" ).
    //
    // Case-insensitive throughout (OrdinalIgnoreCase - these are user-typed part numbers and board
    // labels, not linguistic text, so ordinal is both right and fast).
    //
    // IsEmpty queries match EVERYTHING: an empty search box is not a filter, and returning nothing
    // for it would blank the tab the moment the user cleared the field.
    // ###########################################################################################
    public sealed class WorklogSearchQuery
    {
        private WorklogSearchQuery(IReadOnlyList<WorklogSearchTerm> terms)
        {
            this.Terms = terms;
        }

        public IReadOnlyList<WorklogSearchTerm> Terms { get; }

        // Only the terms that must be PRESENT - the ones worth highlighting. An excluded term never
        // appears in a matched record by definition, so there is nothing of it to highlight.
        public IEnumerable<WorklogSearchTerm> RequiredTerms => this.Terms.Where(t => !t.IsExcluded);

        public bool IsEmpty => this.Terms.Count == 0;

        // ###########################################################################################
        // Parses a raw query string into terms. Never throws and never returns null: malformed input
        // (a lone "-", an unclosed quote, nothing but spaces) yields the terms it could make sense of,
        // because this runs on every keystroke in a search box - refusing to parse would mean the box
        // stops filtering halfway through typing a quoted phrase.
        //
        // An unclosed quote is treated as running to the end of the input, so typing `"black scr`
        // filters on `black scr` rather than on nothing at all.
        // ###########################################################################################
        public static WorklogSearchQuery Parse(string? query)
        {
            var terms = new List<WorklogSearchTerm>();

            if (string.IsNullOrWhiteSpace(query))
                return new WorklogSearchQuery(terms);

            int index = 0;
            while (index < query.Length)
            {
                if (char.IsWhiteSpace(query[index]))
                {
                    index++;
                    continue;
                }

                bool isExcluded = false;
                if (query[index] == '-')
                {
                    isExcluded = true;
                    index++;

                    // A trailing "-" with nothing after it is not a term at all.
                    if (index >= query.Length || char.IsWhiteSpace(query[index]))
                        continue;
                }

                string text;
                if (query[index] == '"')
                {
                    index++;
                    int closing = query.IndexOf('"', index);
                    if (closing < 0)
                    {
                        text = query[index..];
                        index = query.Length;
                    }
                    else
                    {
                        text = query[index..closing];
                        index = closing + 1;
                    }
                }
                else
                {
                    int start = index;
                    while (index < query.Length && !char.IsWhiteSpace(query[index]))
                        index++;

                    text = query[start..index];
                }

                // An empty quoted run ("") carries no constraint - drop it rather than making it a
                // term that matches every string trivially (or, if excluded, matches nothing).
                if (text.Length > 0)
                    terms.Add(new WorklogSearchTerm(text, isExcluded));
            }

            return new WorklogSearchQuery(terms);
        }

        // ###########################################################################################
        // True when the given fields, taken TOGETHER, satisfy the query: every required term appears
        // in at least one of them, and no excluded term appears in any of them.
        //
        // "Together" is the important part - the fields are one record's searchable text, so
        // "cpu super" matches a record whose title says "CPU" and whose comment says "super". Terms
        // are ANDed across the record, not within a single field.
        //
        // Null/blank fields are skipped rather than treated as empty strings that could satisfy a
        // term; an empty query matches everything (see the class header).
        // ###########################################################################################
        public bool Matches(IEnumerable<string?> fields)
        {
            if (this.IsEmpty)
                return true;

            var searchable = fields?.Where(f => !string.IsNullOrEmpty(f)).ToList() ?? new List<string?>();

            foreach (var term in this.Terms)
            {
                bool found = searchable.Any(field =>
                    field!.Contains(term.Text, StringComparison.OrdinalIgnoreCase));

                if (term.IsExcluded && found)
                    return false;

                if (!term.IsExcluded && !found)
                    return false;
            }

            return true;
        }

        public bool Matches(params string?[] fields) => this.Matches((IEnumerable<string?>)fields);

        // ###########################################################################################
        // Every place a REQUIRED term appears in one string, merged and ordered, ready to drive
        // highlighting. Overlapping and adjacent hits are merged into one run so a query like
        // "cpu cp" does not produce two overlapping highlight spans that would double-draw.
        //
        // Returns an empty list for an empty query or a blank string - nothing to highlight, which
        // callers render as plain text.
        // ###########################################################################################
        public IReadOnlyList<WorklogSearchHit> FindHits(string? text)
        {
            var hits = new List<WorklogSearchHit>();

            if (this.IsEmpty || string.IsNullOrEmpty(text))
                return hits;

            foreach (var term in this.RequiredTerms)
            {
                int from = 0;
                while (from <= text.Length - term.Text.Length)
                {
                    int found = text.IndexOf(term.Text, from, StringComparison.OrdinalIgnoreCase);
                    if (found < 0)
                        break;

                    hits.Add(new WorklogSearchHit(found, term.Text.Length));
                    from = found + 1; // +1, not +Length: overlapping occurrences both count.
                }
            }

            if (hits.Count == 0)
                return hits;

            hits.Sort((a, b) => a.Start.CompareTo(b.Start));

            var merged = new List<WorklogSearchHit>();
            int currentStart = hits[0].Start;
            int currentEnd = hits[0].End;

            for (int i = 1; i < hits.Count; i++)
            {
                if (hits[i].Start <= currentEnd)
                {
                    currentEnd = Math.Max(currentEnd, hits[i].End);
                    continue;
                }

                merged.Add(new WorklogSearchHit(currentStart, currentEnd - currentStart));
                currentStart = hits[i].Start;
                currentEnd = hits[i].End;
            }

            merged.Add(new WorklogSearchHit(currentStart, currentEnd - currentStart));
            return merged;
        }

        // ###########################################################################################
        // The same text split into alternating plain and matched segments, in order, covering the
        // WHOLE string with no gaps - ready to be turned into styled runs by a caller that knows how
        // to draw them (see TabWorkbooks.BuildHighlightedTextBlock).
        //
        // The split lives here rather than in the tab because it is where the off-by-one risk is:
        // every segment boundary has to line up exactly or characters get dropped or duplicated on
        // screen. Returning segments rather than Avalonia Inlines keeps it unit-testable, which is
        // the whole reason the maths is on this side of the line.
        //
        // A string with no hits comes back as a single unmatched segment rather than an empty list,
        // so a caller can render the result unconditionally.
        // ###########################################################################################
        public IReadOnlyList<(string Text, bool IsMatch)> SplitIntoSegments(string? text)
        {
            var segments = new List<(string Text, bool IsMatch)>();

            if (string.IsNullOrEmpty(text))
                return segments;

            var hits = this.FindHits(text);
            if (hits.Count == 0)
            {
                segments.Add((text, false));
                return segments;
            }

            int position = 0;
            foreach (var hit in hits)
            {
                if (hit.Start > position)
                    segments.Add((text[position..hit.Start], false));

                segments.Add((text[hit.Start..hit.End], true));
                position = hit.End;
            }

            if (position < text.Length)
                segments.Add((text[position..], false));

            return segments;
        }
    }
}
