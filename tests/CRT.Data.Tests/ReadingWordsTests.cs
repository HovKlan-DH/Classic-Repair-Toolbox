using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ReadingWords - the expected oscilloscope readings in plain words (owner request, 2026-11-17: "convert this in to normal
// words like "Low" and "High" ... make sure it is also done on go-live day"; "FLOAT" and "VIDEO" too, asked).
// ###########################################################################################
public sealed class ReadingWordsTests
{
    [Theory]
    [InlineData("HIGH", "High")]
    [InlineData("LOW", "Low")]
    [InlineData("LOW (or floating?)", "Low (or floating?)")]
    [InlineData("LOW (and floating?)", "Low (and floating?)")]
    [InlineData("LOW(?)", "Low(?)")]
    [InlineData("HIGH(?)", "High(?)")]
    [InlineData("HIGH (or floating?)", "High (or floating?)")]
    [InlineData("FLOAT", "Floating")]
    [InlineData("VIDEO", "Video")]
    // Whatever stands round the word stays, spaces too - as round "HIGH" (code review, 2026-11-17).
    [InlineData(" FLOAT ", " Floating ")]
    [InlineData(" HIGH ", " High ")]
    // FLOAT and VIDEO wherever they stand as a word, as HIGH and LOW are (owner decision, 2026-11-17: "yes, reword those also").
    [InlineData("FLOAT (sometimes)", "Floating (sometimes)")]
    [InlineData("LOW/FLOAT", "Low/Floating")]
    [InlineData("HIGH (FLOAT at reset)", "High (Floating at reset)")]
    [InlineData("VIDEO (composite)", "Video (composite)")]
    // Already in words, or no such word: as it is.
    [InlineData("High", "High")]
    [InlineData("Low", "Low")]
    [InlineData("Floating", "Floating")]
    [InlineData("Pulsing", "Pulsing")]
    [InlineData("0V (GND)", "0V (GND)")]
    [InlineData("1.02272 MHz", "1.02272 MHz")]
    [InlineData("N/A", "N/A")]
    [InlineData("SLOW CLOCK", "SLOW CLOCK")]
    [InlineData("HIGHER", "HIGHER")]
    [InlineData("FLOATING", "FLOATING")]
    [InlineData("VIDEOS", "VIDEOS")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void A_reading_in_capitals_is_written_in_plain_words(string? reading, string words)
    {
        Assert.Equal(words, ReadingWords.Of(reading));
    }

    // Every baseline's reading on the board, each row otherwise as it was; a board with nothing to change is the same board.
    [Fact]
    public void The_whole_board_writes_its_readings_in_plain_words_and_says_how_many()
    {
        BoardData board = new()
        {
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U20", Pin = "1", Name = "1DISCHARGE", ExpectedOscilloscopeReading = "LOW", File = "a.png", TimeDiv = "1ms" },
                new ComponentImageEntry { BoardLabel = "U20", Pin = "5", Name = "1OUTPUT", ExpectedOscilloscopeReading = "Pulsing", File = "b.png" },
                new ComponentImageEntry { BoardLabel = "U7", Pin = "1", Name = "PHI0", ExpectedOscilloscopeReading = "HIGH", File = "c.png" },
                new ComponentImageEntry { BoardLabel = "U7", File = "photo.jpg" },
            ],
        };

        BoardData made = ReadingWords.ApplyToBoard(board, out int rows);

        Assert.Equal(2, rows);
        Assert.Equal(["Low", "Pulsing", "High", ""], made.ComponentImages.Select(image => image.ExpectedOscilloscopeReading));
        Assert.Equal(["a.png", "b.png", "c.png", "photo.jpg"], made.ComponentImages.Select(image => image.File));
        Assert.Equal("1ms", made.ComponentImages[0].TimeDiv);
        Assert.Same(board.ComponentImages[1], made.ComponentImages[1]);

        Assert.Same(made, ReadingWords.ApplyToBoard(made, out int again));
        Assert.Equal(0, again);
    }
}
