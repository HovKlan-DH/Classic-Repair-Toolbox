using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // What a KiCad view is: a side of a board, or a schematic sheet.
    public enum KiCadViewSide
    {
        Sheet,
        Top,
        Bottom,
    }

    // ###########################################################################################
    // A KiCad view as matching sees it: its name - the CAD name a picture takes - what it is, and for
    // a side of a board the name of the PCB file it comes from ("Issue 4" for "Issue 4 - PCB Top").
    // ###########################################################################################
    public sealed record KiCadViewToMatch(string Name, KiCadViewSide Side, string Source = "");

    // One suggested match: the picture, by its row's name and image file, and the view it seems to show.
    public sealed record SchematicCadNameSuggestion(string SchematicName, string ImageFile, string CadName);

    // ###########################################################################################
    // WHICH KICAD VIEW EACH PICTURE SEEMS TO SHOW (CRT 4.0 step 3, slice 3d - the plan: "Match:
    // suggested pairs ("...Top.png" with "PCB Top") to confirm").
    //
    // Only a picture with an image and no CAD name is matched, and only to a view no picture has. Its
    // NAME and its FILE's name are read as words (capitals ignored, "&" read as "and"), and a pair is
    // scored by what the words say - the higher the surer:
    //
    //   - THE SAME WORDS as the view's name, the picture's name or its file's ("Clock" and "Clock",
    //     "C64-250469-KiCad.png" and "C64-250469-KiCad").
    //   - A SIDE OF THE BOARD said in words - top or front, bottom, back or solder - with the other
    //     side not said, for a picture that is no schematic: that side's view, surer when the words
    //     also name its PCB file. Then KiCad's own plot spellings - "Issue 4.F.png", "Issue 4.B.png",
    //     "board-F_Cu" - never a lone letter anywhere in a name ("C64 Rev B" is no bottom).
    //   - A SHEET'S NAME ending the picture's words - KiCad plots a sheet as "<project>-<parent>-
    //     <sheet>.png", so the last words are the sheet ("Open128-I_O-Control Ports.png" is "Control
    //     Ports", not "I/O") - surer the more words it has; anywhere else in them, much less sure.
    //
    // The surest pairs are taken first. *** WHERE TWO PAIRS ARE EQUALLY SURE OF ONE PICTURE OR ONE
    // VIEW, NEITHER IS SUGGESTED *** - nor anything else for either of them: "250407_", the C64's
    // first sheet, is a word of three of its pictures' files, and a guess would be wrong twice out of
    // three. Nothing suggested is better than a wrong suggestion a contributor confirms without
    // looking. Measured on the five live boards with KiCad data, their CAD names taken away
    // (2026-10-28, SchematicCadNameSuggestionsTests): 34 suggestions, every one the CAD name the board
    // has; 250407's two schematic sheets are the only pictures with a view left unsuggested.
    //
    // PURE: the rows and the views in, the pairs out, in the rows' order.
    // ###########################################################################################
    public static class SchematicCadNameSuggestions
    {
        private const int SameWords = 1000;
        private const int SideInWords = 300;
        private const int SheetAtEnd = 200;
        private const int SideInLetters = 150;
        private const int NamesItsPcb = 20;
        private const int PerSheetWord = 10;

        private static readonly HashSet<string> TopWords = new(StringComparer.Ordinal) { "top", "front" };
        private static readonly HashSet<string> BottomWords = new(StringComparer.Ordinal) { "bottom", "back", "solder", "underside" };

        public static List<SchematicCadNameSuggestion> For(IReadOnlyList<BoardSchematicEntry> schematics, IReadOnlyList<KiCadViewToMatch> views)
        {
            ArgumentNullException.ThrowIfNull(schematics);
            ArgumentNullException.ThrowIfNull(views);

            var held = new HashSet<string>(
                schematics.Select(row => row.CadName?.Trim() ?? string.Empty).Where(name => name.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            // The pictures to match: with an image, no CAD name, and a name and file no other row
            // shares - a suggestion finds its row by both.
            List<(int Row, Picture Picture)> pictures = schematics
                .Select((row, index) => (Row: index, Entry: row))
                .Where(item => !string.IsNullOrWhiteSpace(item.Entry.SchematicImageFile) && string.IsNullOrWhiteSpace(item.Entry.CadName))
                .Where(item => schematics.Count(other =>
                    SchematicRename.SameName(other.SchematicName, item.Entry.SchematicName) &&
                    SchematicRows.SameImageFile(other.SchematicImageFile, item.Entry.SchematicImageFile)) == 1)
                .Select(item => (item.Row, Picture.Of(item.Entry)))
                .ToList();

            List<KiCadViewToMatch> free = views
                .Select(view => view with { Name = (view.Name ?? string.Empty).Trim() })
                .Where(view => view.Name.Length > 0 && !held.Contains(view.Name))
                .GroupBy(view => view.Name, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

            var candidates = new List<(int Picture, int View, int Score)>();

            for (int p = 0; p < pictures.Count; p++)
            {
                for (int v = 0; v < free.Count; v++)
                {
                    int score = SchematicCadNameSuggestions.Score(pictures[p].Picture, free[v]);

                    if (score > 0)
                    {
                        candidates.Add((p, v, score));
                    }
                }
            }

            var takenPictures = new Dictionary<int, int>();
            var takenViews = new HashSet<int>();
            var contestedPictures = new HashSet<int>();
            var contestedViews = new HashSet<int>();

            foreach (IGrouping<int, (int Picture, int View, int Score)> level in candidates.GroupBy(c => c.Score).OrderByDescending(group => group.Key))
            {
                List<(int Picture, int View, int Score)> open = level
                    .Where(c => !takenPictures.ContainsKey(c.Picture) && !takenViews.Contains(c.View) &&
                                !contestedPictures.Contains(c.Picture) && !contestedViews.Contains(c.View))
                    .ToList();

                foreach ((int picture, int view, _) in open)
                {
                    bool alone = open.Count(c => c.Picture == picture) == 1 && open.Count(c => c.View == view) == 1;

                    if (alone)
                    {
                        takenPictures[picture] = view;
                        takenViews.Add(view);
                    }
                    else
                    {
                        contestedPictures.Add(picture);
                        contestedViews.Add(view);
                    }
                }
            }

            return takenPictures
                .OrderBy(pair => pictures[pair.Key].Row)
                .Select(pair =>
                {
                    BoardSchematicEntry row = schematics[pictures[pair.Key].Row];
                    return new SchematicCadNameSuggestion(row.SchematicName?.Trim() ?? string.Empty, row.SchematicImageFile?.Trim() ?? string.Empty, free[pair.Value].Name);
                })
                .ToList();
        }

        // ###########################################################################################
        // The board with the suggestions taken - each one only while its picture still has no CAD
        // name and no picture has its view (the draft can have moved since they were worked out) -
        // and how many were.
        // ###########################################################################################
        public static BoardData With(BoardData board, IReadOnlyList<SchematicCadNameSuggestion> suggestions, out IReadOnlyList<SchematicCadNameSuggestion> taken)
        {
            BoardData result = SchematicCadNameSuggestions.Take(board, suggestions, out List<SchematicCadNameSuggestion> took);
            taken = took;
            return result;
        }

        // The suggestions the board can still take, in their order - what With would take.
        public static List<SchematicCadNameSuggestion> StillOpen(BoardData board, IReadOnlyList<SchematicCadNameSuggestion> suggestions)
        {
            SchematicCadNameSuggestions.Take(board, suggestions, out List<SchematicCadNameSuggestion> took);
            return took;
        }

        private static BoardData Take(BoardData board, IReadOnlyList<SchematicCadNameSuggestion> suggestions, out List<SchematicCadNameSuggestion> taken)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(suggestions);

            taken = [];
            BoardData result = board;

            foreach (SchematicCadNameSuggestion suggestion in suggestions)
            {
                if (!SchematicCadNameSuggestions.IsOpen(result, suggestion))
                {
                    continue;
                }

                (result, bool found) = SchematicRows.WithCadName(result, new SchematicRowRef(suggestion.SchematicName, suggestion.ImageFile), suggestion.CadName);

                if (found)
                {
                    taken.Add(suggestion);
                }
            }

            return result;
        }

        private static bool IsOpen(BoardData board, SchematicCadNameSuggestion suggestion)
        {
            int index = SchematicRows.IndexOf(board.Schematics, new SchematicRowRef(suggestion.SchematicName, suggestion.ImageFile));

            return index >= 0 &&
                   string.IsNullOrWhiteSpace(board.Schematics[index].CadName) &&
                   !board.Schematics.Any(row => string.Equals(row.CadName?.Trim(), suggestion.CadName.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        // ---- the scores -----------------------------------------------------------------------------

        // A picture read as words - its name's, and its file's name's without the extension - and the
        // two as written, where KiCad's plot letters are looked for (PlotSideOf).
        private sealed record Picture(IReadOnlyList<string> Name, IReadOnlyList<string> File, string WrittenName, string WrittenFile)
        {
            public static Picture Of(BoardSchematicEntry row)
            {
                string file = (row.SchematicImageFile ?? string.Empty).Trim().Replace('\\', '/');
                string fileName = file[(file.LastIndexOf('/') + 1)..];
                int dot = fileName.LastIndexOf('.');
                string stem = dot > 0 ? fileName[..dot] : fileName;

                return new Picture(
                    SchematicCadNameSuggestions.Words(row.SchematicName),
                    SchematicCadNameSuggestions.Words(stem),
                    row.SchematicName?.Trim() ?? string.Empty,
                    stem);
            }

            public IEnumerable<string> All => this.Name.Concat(this.File);
        }

        // ###########################################################################################
        // A side of the board as KiCad writes it: a plot named "<name>.F" or "<name>.B", or the copper
        // layer "F.Cu" / "B_Cu". Never a lone letter anywhere in a name (code review, 2026-10-28): "C64
        // Rev B" and "BBC Model B" read as the bottom of the board - and never a revision or model
        // letter after a dot (second code review, 2026-10-28): "250407 Rev.B.jpg" did too.
        // ###########################################################################################
        private const string NotARevision = @"(?<!(^|[^a-z0-9])(rev|revision|ver|version|model|issue|mk|mark))";

        private static readonly Regex PlotTop = new(NotARevision + @"\.f$|(^|[^a-z0-9])f[._]cu($|[^a-z0-9])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex PlotBottom = new(NotARevision + @"\.b$|(^|[^a-z0-9])b[._]cu($|[^a-z0-9])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static KiCadViewSide? PlotSideOf(Picture picture) =>
            SchematicCadNameSuggestions.SideOf(
                PlotTop.IsMatch(picture.WrittenName) || PlotTop.IsMatch(picture.WrittenFile),
                PlotBottom.IsMatch(picture.WrittenName) || PlotBottom.IsMatch(picture.WrittenFile));

        private static int Score(Picture picture, KiCadViewToMatch view)
        {
            List<string> viewWords = SchematicCadNameSuggestions.Words(view.Name);

            if (viewWords.Count == 0)
            {
                return 0;
            }

            if (viewWords.SequenceEqual(picture.Name) || viewWords.SequenceEqual(picture.File))
            {
                return SchematicCadNameSuggestions.SameWords;
            }

            return view.Side == KiCadViewSide.Sheet
                ? SchematicCadNameSuggestions.SheetScore(picture, viewWords)
                : SchematicCadNameSuggestions.SideScore(picture, view);
        }

        // A sheet's name among the picture's words - at their end, or anywhere.
        private static int SheetScore(Picture picture, List<string> viewWords)
        {
            int best = 0;

            foreach (IReadOnlyList<string> words in new[] { picture.Name, picture.File })
            {
                int at = SchematicCadNameSuggestions.RunAt(words, viewWords, fromEnd: true);

                if (at >= 0)
                {
                    best = Math.Max(best, SchematicCadNameSuggestions.SheetAtEnd + SchematicCadNameSuggestions.PerSheetWord * viewWords.Count);
                }
                else if (SchematicCadNameSuggestions.RunAt(words, viewWords, fromEnd: false) >= 0)
                {
                    best = Math.Max(best, SchematicCadNameSuggestions.PerSheetWord * viewWords.Count);
                }
            }

            return best;
        }

        // A side of the board said by a picture that is no schematic: in words, else in KiCad's letters.
        private static int SideScore(Picture picture, KiCadViewToMatch view)
        {
            List<string> all = picture.All.ToList();

            if (all.Any(word => word is "schematic" or "schematics" or "page" || word.StartsWith("sheet", StringComparison.Ordinal)))
            {
                return 0;
            }

            KiCadViewSide? inWords = SchematicCadNameSuggestions.SideOf(all.Any(TopWords.Contains), all.Any(BottomWords.Contains));
            KiCadViewSide? inLetters = SchematicCadNameSuggestions.PlotSideOf(picture);

            int score = inWords == view.Side
                ? SchematicCadNameSuggestions.SideInWords
                : inWords is null && inLetters == view.Side
                    ? SchematicCadNameSuggestions.SideInLetters
                    : 0;

            List<string> source = SchematicCadNameSuggestions.Words(view.Source);

            if (score > 0 && source.Count > 0 &&
                (SchematicCadNameSuggestions.RunAt(picture.Name, source, fromEnd: false) >= 0 || SchematicCadNameSuggestions.RunAt(picture.File, source, fromEnd: false) >= 0))
            {
                score += SchematicCadNameSuggestions.NamesItsPcb;
            }

            return score;
        }

        // One side said and not the other.
        private static KiCadViewSide? SideOf(bool top, bool bottom) =>
            top == bottom ? null : top ? KiCadViewSide.Top : KiCadViewSide.Bottom;

        // Where `run` stands in `words` as one stretch - at their very end only, or anywhere - or -1.
        private static int RunAt(IReadOnlyList<string> words, IReadOnlyList<string> run, bool fromEnd)
        {
            if (run.Count == 0 || run.Count > words.Count)
            {
                return -1;
            }

            for (int start = fromEnd ? words.Count - run.Count : 0; start <= words.Count - run.Count; start++)
            {
                bool same = true;

                for (int i = 0; i < run.Count && same; i++)
                {
                    same = string.Equals(words[start + i], run[i], StringComparison.Ordinal);
                }

                if (same)
                {
                    return start;
                }
            }

            return -1;
        }

        // A name as its words: letters and digits, in small letters, "&" read as "and".
        private static List<string> Words(string? text)
        {
            var words = new List<string>();
            var word = new System.Text.StringBuilder();

            foreach (char c in (text ?? string.Empty).Replace("&", " and ", StringComparison.Ordinal))
            {
                if (char.IsLetterOrDigit(c))
                {
                    word.Append(char.ToLowerInvariant(c));
                }
                else if (word.Length > 0)
                {
                    words.Add(word.ToString());
                    word.Clear();
                }
            }

            if (word.Length > 0)
            {
                words.Add(word.ToString());
            }

            return words;
        }
    }
}
