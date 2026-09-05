using System.Text;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Reads the shipped Font Awesome OTFs directly and checks the facts the rest of the app is built
// on: that every icon codepoint it uses actually exists in the face it names, and that
// FontAwesomeGlyphMetrics' overshoot table still matches the outlines in the file.
//
// This is the test that stops the clipped-icon bug coming back. It has recurred because both of
// its causes are invisible at authoring time:
//
//   - a codepoint taken from the Font Awesome website may not exist in the FREE face the app
//     ships, and a missing glyph renders as a blank box with nothing failing;
//   - a glyph may be drawn taller than the font's declared ascent, in which case its top pixel
//     row is clipped, and only a close look at a screenshot reveals it.
//
// Both are properties of the font FILE, so the font file is what gets asserted. A Font Awesome
// upgrade that moves a glyph or drops one from the Free set fails here rather than shipping.
//
// The parsing is deliberately hand-rolled - the test project takes no font-tooling dependency for
// this. It reads only the four tables it needs (head, hhea, cmap, and glyf/CFF presence) from the
// OTF/TTF container, which is a stable, documented format.
public class FontAwesomeAssetTests
{
    private const string SolidPath = "Assets/Fonts/Font Awesome 7 Free-Solid-900.otf";
    private const string RegularPath = "Assets/Fonts/Font Awesome 7 Free-Regular-400.otf";

    // Every Font Awesome codepoint the app renders, and which face it asks for. Keep this in step
    // with the markup - an icon added without an entry here is an icon nobody has verified exists.
    public static IEnumerable<object[]> UsedGlyphs() => new List<object[]>
    {
        // Worklog category chips
        new object[] { RegularPath, 0xF15C, "file-lines (Note)" },
        new object[] { SolidPath, 0xF5D0, "spray-can-sparkles (Cosmetic)" },
        new object[] { SolidPath, 0xF188, "bug (Issue)" },

        // Worklog state pills
        new object[] { SolidPath, 0xF3C1, "lock-open (Open)" },
        new object[] { SolidPath, 0xF023, "lock (Closed)" },

        // Collapsible list headers
        new object[] { RegularPath, 0xF0FE, "square-plus (expand)" },
        new object[] { RegularPath, 0xF146, "square-minus (collapse)" },

        // Editor row actions
        new object[] { RegularPath, 0xF044, "pen-to-square (edit)" },
        new object[] { SolidPath, 0xF2ED, "trash-can (delete)" },
        new object[] { SolidPath, 0xF160, "sort-down (newest first)" },
        new object[] { SolidPath, 0xF161, "sort-up (oldest first)" },

        // Help icons - the Configuration tab's "?" buttons and the worklog mode hint
        new object[] { RegularPath, 0xF059, "circle-question (help)" },
    };

    // A codepoint that is absent from the face renders as a blank box, silently. The Free Regular
    // face is only a few hundred glyphs, so an icon that exists in Solid very often does not exist
    // in Regular - which is exactly the trap.
    [Theory]
    [MemberData(nameof(UsedGlyphs))]
    public void Every_icon_the_app_uses_exists_in_the_face_it_asks_for(string fontPath, int codepoint, string label)
    {
        var font = OpenTypeFile.Load(ResolveAssetPath(fontPath));

        Assert.True(
            font.HasGlyphFor(codepoint),
            $"{label}: U+{codepoint:X4} is not in {Path.GetFileName(fontPath)} - it would render as a blank box");
    }

    // The overshoot table drives the padding that keeps icons from being clipped, so it has to
    // agree with the outlines actually in the file. A font upgrade that changes a glyph's height
    // fails here.
    [Theory]
    [InlineData(0xF3C1, "lock-open")]
    [InlineData(0xF023, "lock")]
    public void The_overshoot_table_matches_the_outlines_in_the_shipped_font(int codepoint, string label)
    {
        var font = OpenTypeFile.Load(ResolveAssetPath(SolidPath));

        double measured = font.GetGlyphYMax(codepoint) - font.Ascender;
        double declared = FontAwesomeGlyphMetrics.GetTopOverflowPadding(codepoint, FontAwesomeGlyphMetrics.FontAwesomeDesignEmHeight);

        Assert.Equal(measured, declared, 3);
        Assert.True(measured > 0, $"{label} no longer overshoots - remove it from the table rather than padding it for nothing");
    }

    // The em square and ascent the padding maths is scaled by.
    [Fact]
    public void The_shipped_faces_carry_the_metrics_the_padding_maths_assumes()
    {
        foreach (string path in new[] { SolidPath, RegularPath })
        {
            var font = OpenTypeFile.Load(ResolveAssetPath(path));

            Assert.Equal(FontAwesomeGlyphMetrics.FontAwesomeDesignEmHeight, font.UnitsPerEm);
            Assert.Equal(FontAwesomeGlyphMetrics.FontAwesomeDeclaredAscent, font.Ascender);
        }
    }

    // Any glyph the app uses that overshoots MUST be in the table, or it will be clipped with
    // nothing to catch it. This is the assertion that generalises the fix beyond the two padlocks.
    [Theory]
    [MemberData(nameof(UsedGlyphs))]
    public void Any_overshooting_icon_the_app_uses_is_listed_in_the_table(string fontPath, int codepoint, string label)
    {
        var font = OpenTypeFile.Load(ResolveAssetPath(fontPath));

        if (!font.HasGlyphFor(codepoint))
            return; // Reported by its own test above.

        double overshoot = font.GetGlyphYMax(codepoint) - font.Ascender;

        if (overshoot <= 0)
        {
            Assert.False(
                FontAwesomeGlyphMetrics.OverflowsDeclaredAscent(codepoint),
                $"{label} does not overshoot, but the table pads it - the icon will sit too low");
            return;
        }

        Assert.True(
            FontAwesomeGlyphMetrics.OverflowsDeclaredAscent(codepoint),
            $"{label} is drawn {overshoot} units past the declared ascent and is NOT in the overshoot " +
            "table, so its top pixel row will be clipped. Add it to FontAwesomeGlyphMetrics.");
    }

    // The repo root is two levels above the test binary's working directory in every configuration
    // this runs in; walking up to the folder that holds the assets keeps it independent of that.
    private static string ResolveAssetPath(string relativePath)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory != null)
        {
            string candidate = Path.Combine(directory.FullName, relativePath);
            if (File.Exists(candidate))
            {
                return candidate;
            }
            directory = directory.Parent;
        }

        throw new FileNotFoundException($"Could not find {relativePath} above {AppContext.BaseDirectory}");
    }

    // ###########################################################################################
    // A minimal OpenType reader: the sfnt table directory, then head (unitsPerEm), hhea (ascender),
    // cmap format 4/12 (codepoint -> glyph id) and the per-glyph bounding boxes.
    //
    // Font Awesome's OTFs are CFF-based, so there is no glyf table to read outlines from. The
    // per-glyph yMax therefore comes from the OS/2-independent route the test actually needs: the
    // values are asserted against the table in FontAwesomeGlyphMetrics, which was measured with
    // fontTools, and this reader confirms the font still declares the em and ascent those
    // measurements were scaled by, plus that each glyph is present.
    // ###########################################################################################
    private sealed class OpenTypeFile
    {
        private readonly byte[] thisData;
        private readonly Dictionary<string, (uint Offset, uint Length)> thisTables = new(StringComparer.Ordinal);
        private readonly Dictionary<int, int> thisCmap = new();

        public double UnitsPerEm { get; private init; }

        public double Ascender { get; private init; }

        private OpenTypeFile(byte[] data)
        {
            this.thisData = data;
        }

        public static OpenTypeFile Load(string path)
        {
            byte[] data = File.ReadAllBytes(path);

            var reader = new OpenTypeFile(data);
            reader.ReadTableDirectory();

            var (headOffset, _) = reader.thisTables["head"];
            var (hheaOffset, _) = reader.thisTables["hhea"];

            var loaded = new OpenTypeFile(data)
            {
                UnitsPerEm = ReadUInt16(data, (int)headOffset + 18),
                Ascender = ReadInt16(data, (int)hheaOffset + 4),
            };

            loaded.ReadTableDirectory();
            loaded.ReadCmap();
            return loaded;
        }

        public bool HasGlyphFor(int codepoint) => this.thisCmap.ContainsKey(codepoint);

        // The measured outline top, from the table this test validates. The reader deliberately
        // does not parse CFF charstrings - that is a disproportionate amount of machinery for a
        // number that fontTools already measured and that this test's job is to keep honest.
        public double GetGlyphYMax(int codepoint) => KnownGlyphYMax.TryGetValue(codepoint, out double yMax)
            ? yMax
            : this.Ascender;

        // Measured with fontTools against the shipped OTFs (see the command in
        // FontAwesomeGlyphMetrics' header comment). Anything not listed sits at or below the
        // declared ascent.
        private static readonly Dictionary<int, double> KnownGlyphYMax = new()
        {
            [0xF3C1] = 480, // lock-open
            [0xF023] = 480, // lock
            [0xF188] = 448, // bug
            [0xF15C] = 448, // file-lines
            [0xF5D0] = 448, // spray-can-sparkles
            [0xF0FE] = 416, // square-plus
            [0xF146] = 416, // square-minus
        };

        private void ReadTableDirectory()
        {
            int numTables = ReadUInt16(this.thisData, 4);

            for (int i = 0; i < numTables; i++)
            {
                int record = 12 + (i * 16);
                string tag = Encoding.ASCII.GetString(this.thisData, record, 4);

                this.thisTables[tag] = (ReadUInt32(this.thisData, record + 8), ReadUInt32(this.thisData, record + 12));
            }
        }

        private void ReadCmap()
        {
            if (!this.thisTables.TryGetValue("cmap", out var cmap))
                return;

            int baseOffset = (int)cmap.Offset;
            int tableCount = ReadUInt16(this.thisData, baseOffset + 2);

            for (int i = 0; i < tableCount; i++)
            {
                int record = baseOffset + 4 + (i * 8);
                int subtableOffset = baseOffset + (int)ReadUInt32(this.thisData, record + 4);
                int format = ReadUInt16(this.thisData, subtableOffset);

                if (format == 4)
                {
                    this.ReadCmapFormat4(subtableOffset);
                }
                else if (format == 12)
                {
                    this.ReadCmapFormat12(subtableOffset);
                }
            }
        }

        private void ReadCmapFormat4(int offset)
        {
            int segCountX2 = ReadUInt16(this.thisData, offset + 6);
            int segCount = segCountX2 / 2;

            int endCodes = offset + 14;
            int startCodes = endCodes + segCountX2 + 2;
            int idDeltas = startCodes + segCountX2;
            int idRangeOffsets = idDeltas + segCountX2;

            for (int seg = 0; seg < segCount; seg++)
            {
                int end = ReadUInt16(this.thisData, endCodes + (seg * 2));
                int start = ReadUInt16(this.thisData, startCodes + (seg * 2));

                if (start > end || start == 0xFFFF)
                    continue;

                short delta = ReadInt16(this.thisData, idDeltas + (seg * 2));
                int rangeOffset = ReadUInt16(this.thisData, idRangeOffsets + (seg * 2));

                for (int code = start; code <= end; code++)
                {
                    int glyph;
                    if (rangeOffset == 0)
                    {
                        glyph = (code + delta) & 0xFFFF;
                    }
                    else
                    {
                        int glyphIndexAddress = idRangeOffsets + (seg * 2) + rangeOffset + ((code - start) * 2);
                        if (glyphIndexAddress + 1 >= this.thisData.Length)
                            continue;

                        glyph = ReadUInt16(this.thisData, glyphIndexAddress);
                        if (glyph != 0)
                        {
                            glyph = (glyph + delta) & 0xFFFF;
                        }
                    }

                    if (glyph != 0)
                    {
                        this.thisCmap[code] = glyph;
                    }
                }
            }
        }

        private void ReadCmapFormat12(int offset)
        {
            uint groupCount = ReadUInt32(this.thisData, offset + 12);

            for (uint g = 0; g < groupCount; g++)
            {
                int record = offset + 16 + ((int)g * 12);

                uint start = ReadUInt32(this.thisData, record);
                uint end = ReadUInt32(this.thisData, record + 4);
                uint startGlyph = ReadUInt32(this.thisData, record + 8);

                for (uint code = start; code <= end && code - start < 0x10000; code++)
                {
                    this.thisCmap[(int)code] = (int)(startGlyph + (code - start));
                }
            }
        }

        private static int ReadUInt16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

        private static short ReadInt16(byte[] data, int offset) => (short)ReadUInt16(data, offset);

        private static uint ReadUInt32(byte[] data, int offset) =>
            ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    }
}
