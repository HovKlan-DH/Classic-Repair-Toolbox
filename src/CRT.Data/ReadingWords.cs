using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE EXPECTED READINGS IN PLAIN WORDS (owner request, 2026-11-17: "I have a lot of values for the "Expected oscilloscope
    // reading" where it is set to either "LOW" or "HIGH", but I would actually like to convert this in to normal words like
    // "Low" and "High" - can you do this, and make sure it is also done on go-live day?"; and, asked, "FLOAT" and "VIDEO" too):
    // a scope baseline's expected reading written in capitals is written as a word - "HIGH" is "High" and "LOW" "Low", "FLOAT"
    // "Floating" and "VIDEO" "Video", the spellings the boards already use beside them - each wherever it stands as a word of
    // its own ("LOW (or floating?)" is "Low (or floating?)", "HIGH(?)" "High(?)", "LOW/FLOAT" "Low/Floating"; owner decision,
    // the same day: "yes, reword those also" - FLOAT and VIDEO only as the whole reading until then). Measured on the published
    // boards that day: 2,830 of the 13,862 readings, every other one ("Pulsing", "0V (GND)", "1.02272 MHz") left as it is.
    //
    // DataTreeConversion runs it on every board, so release day's conversion does it again; an Excel-era draft converted
    // (DraftConversion) gets it in the draft and its base, or each of its baselines would read as changed against the
    // published board. Nothing else writes it: a reading typed in CRT stays as typed.
    // ###########################################################################################
    public static class ReadingWords
    {
        private static readonly Regex InCapitals = new(@"\b(HIGH|LOW|FLOAT|VIDEO)\b", RegexOptions.CultureInvariant);

        // `reading` in plain words - the same text when it holds none of the words in capitals. Whatever stands round each word
        // stays, spaces too.
        public static string Of(string? reading) =>
            ReadingWords.InCapitals.Replace(reading ?? string.Empty, match => match.Value switch
            {
                "HIGH" => "High",
                "LOW" => "Low",
                "FLOAT" => "Floating",
                _ => "Video",
            });

        // The board with every baseline's reading in plain words - the same board when none changes. `rows` says how many did.
        public static BoardData ApplyToBoard(BoardData board, out int rows)
        {
            ArgumentNullException.ThrowIfNull(board);

            int changed = 0;

            List<ComponentImageEntry> images = board.ComponentImages.Select(row =>
            {
                string reading = ReadingWords.Of(row.ExpectedOscilloscopeReading);

                if (string.Equals(reading, row.ExpectedOscilloscopeReading, StringComparison.Ordinal))
                {
                    return row;
                }

                changed++;
                return row.WithExpectedReading(reading);
            }).ToList();

            rows = changed;

            return changed == 0 ? board : board.WithComponentSections(componentImages: images);
        }

        // The board with every baseline's reading in plain words.
        public static BoardData ApplyToBoard(BoardData board) => ReadingWords.ApplyToBoard(board, out _);
    }
}
