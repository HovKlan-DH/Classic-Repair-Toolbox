using System;
using System.Linq;
using System.Text;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionContentRules - whether a submitted file's opening bytes are what its name
    // claims (security review, 2026-09-25).
    //
    // The name is chosen by the contributor; the bytes are what a user's machine opens. Each test
    // below is one way those two could disagree and have to be refused, or one legitimate shape a
    // real board file takes that must not be.
    // ###########################################################################################
    public sealed class SubmissionContentRulesTests
    {
        private static readonly byte[] Png = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13];

        private static readonly byte[] Jpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0, 16];

        private static readonly byte[] WindowsExecutable = [(byte)'M', (byte)'Z', 0x90, 0, 3, 0, 0, 0];

        private static bool Matches(string path, byte[] head) =>
            SubmissionContentRules.Matches(path, head, out _);

        [Fact]
        public void Each_allowed_type_passes_with_its_real_signature()
        {
            Assert.True(SubmissionContentRulesTests.Matches("a.png", SubmissionContentRulesTests.Png));
            Assert.True(SubmissionContentRulesTests.Matches("a.jpg", SubmissionContentRulesTests.Jpeg));
            Assert.True(SubmissionContentRulesTests.Matches("a.jpeg", SubmissionContentRulesTests.Jpeg));
            Assert.True(SubmissionContentRulesTests.Matches("a.gif", Encoding.ASCII.GetBytes("GIF89a....")));
            Assert.True(SubmissionContentRulesTests.Matches("a.gif", Encoding.ASCII.GetBytes("GIF87a....")));
            Assert.True(SubmissionContentRulesTests.Matches("a.bmp", Encoding.ASCII.GetBytes("BM" + new string('x', 30))));
            Assert.True(SubmissionContentRulesTests.Matches("a.webp", Encoding.ASCII.GetBytes("RIFF\u0001\u0002\u0003\u0004WEBPVP8 ")));
            Assert.True(SubmissionContentRulesTests.Matches("a.pdf", Encoding.ASCII.GetBytes("%PDF-1.7\n")));
            Assert.True(SubmissionContentRulesTests.Matches("a.txt", Encoding.UTF8.GetBytes("Pin 3 reads 5.0V (\u00b1 5%)")));
            Assert.True(SubmissionContentRulesTests.Matches("a.html", Encoding.UTF8.GetBytes("<!DOCTYPE html><html></html>")));
        }

        // The extension is read case-insensitively, as everywhere else it decides a file TYPE -
        // the shipped tree carries .PDF and .JPG alongside .pdf and .jpg.
        [Fact]
        public void An_upper_case_extension_is_the_same_type()
        {
            Assert.True(SubmissionContentRulesTests.Matches("SCAN.JPG", SubmissionContentRulesTests.Jpeg));
            Assert.True(SubmissionContentRulesTests.Matches("Manual.PDF", Encoding.ASCII.GetBytes("%PDF-1.4")));
        }

        // THE attack this class exists for: a program renamed to something a user would open.
        [Theory]
        [InlineData("board.png")]
        [InlineData("board.jpg")]
        [InlineData("manual.pdf")]
        [InlineData("notes.txt")]
        [InlineData("page.html")]
        public void An_executable_wearing_a_document_name_is_refused(string path)
        {
            Assert.False(SubmissionContentRulesTests.Matches(path, SubmissionContentRulesTests.WindowsExecutable));
        }

        // ###########################################################################################
        // *** AN IMAGE NAMED AS ANOTHER IMAGE FORMAT IS ACCEPTED, and that was learned from data. ***
        //
        // The first version refused it, and SubmissionRulesShippedDataTests failed on three images
        // every Commodore board cites: a PNG saved as ".jpg" and two JPEGs saved as ".png". They
        // decode fine everywhere, by content. Refusing them made every Commodore board
        // unsubmittable - the third time this pipeline would have rejected published data.
        // ###########################################################################################
        [Fact]
        public void An_image_of_one_format_named_as_another_is_accepted()
        {
            Assert.True(SubmissionContentRulesTests.Matches("photo.jpg", SubmissionContentRulesTests.Png));
            Assert.True(SubmissionContentRulesTests.Matches("photo.png", SubmissionContentRulesTests.Jpeg));
        }

        // ...but an image NAME must still carry image BYTES. A PDF or text named ".png" is refused.
        [Fact]
        public void A_non_image_named_as_an_image_is_refused()
        {
            Assert.False(SubmissionContentRulesTests.Matches("photo.png", Encoding.ASCII.GetBytes("%PDF-1.7")));
            Assert.False(SubmissionContentRulesTests.Matches("photo.jpg", Encoding.ASCII.GetBytes("<html>")));
        }

        // Readers accept a PDF whose header follows some leading junk, within the first kilobyte -
        // and real scanned manuals do this. Refusing them would reject real board documents.
        [Fact]
        public void A_pdf_header_after_leading_bytes_is_accepted_within_the_first_kilobyte()
        {
            byte[] late = Enumerable.Repeat((byte)' ', 500).Concat(Encoding.ASCII.GetBytes("%PDF-1.3")).ToArray();
            byte[] tooLate = Enumerable.Repeat((byte)' ', 2000).Concat(Encoding.ASCII.GetBytes("%PDF-1.3")).ToArray();

            Assert.True(SubmissionContentRulesTests.Matches("a.pdf", late));
            Assert.False(SubmissionContentRulesTests.Matches("a.pdf", tooLate));
        }

        // Text is text: any NUL refuses it. That one test covers executables, archives, images and
        // UTF-16 alike - and the message says how to save it instead.
        [Fact]
        public void A_text_file_carrying_a_NUL_byte_is_refused_with_advice()
        {
            byte[] utf16 = Encoding.Unicode.GetBytes("hello");

            Assert.False(SubmissionContentRules.Matches("notes.txt", utf16, out string reason));
            Assert.Contains("UTF-8", reason, StringComparison.Ordinal);
        }

        // An empty text file is still text; an empty image is not an image.
        [Fact]
        public void An_empty_file_is_text_but_never_an_image()
        {
            Assert.True(SubmissionContentRulesTests.Matches("empty.txt", []));
            Assert.False(SubmissionContentRulesTests.Matches("empty.png", []));
            Assert.False(SubmissionContentRulesTests.Matches("empty.pdf", []));
        }

        // An extension this class does not know fails CLOSED - SubmissionFileRules should never let
        // one through, and if the two lists drift, a file of unknown type must not pass on that.
        [Theory]
        [InlineData("a.svg")]
        [InlineData("a.exe")]
        [InlineData("a.json")]
        [InlineData("a")]
        public void An_unknown_type_never_matches(string path)
        {
            Assert.False(SubmissionContentRulesTests.Matches(path, Encoding.ASCII.GetBytes("<svg></svg>")));
        }

        // The allowlist and this class must know the same types - a type added to one only would
        // either refuse everything of that type (here) or let it through unchecked.
        [Fact]
        public void Every_allowed_extension_has_a_content_rule()
        {
            foreach (string extension in SubmissionFileRules.AllowedExtensions)
            {
                string path = "file" + extension;
                byte[] sample = extension switch
                {
                    ".png" => SubmissionContentRulesTests.Png,
                    ".jpg" or ".jpeg" => SubmissionContentRulesTests.Jpeg,
                    ".gif" => Encoding.ASCII.GetBytes("GIF89a"),
                    ".bmp" => Encoding.ASCII.GetBytes("BM" + new string('x', 30)),
                    ".webp" => Encoding.ASCII.GetBytes("RIFF1234WEBP"),
                    ".pdf" => Encoding.ASCII.GetBytes("%PDF-1.7"),
                    _ => Encoding.ASCII.GetBytes("plain text")
                };

                Assert.True(
                    SubmissionContentRulesTests.Matches(path, sample),
                    $"{extension} is allowed but has no content rule that accepts a real file of it.");
            }
        }

        // ---------------------------------------------------------------------- KiCad data (2026-09-26)

        // ###########################################################################################
        // A KiCad file is plain text opening with its format's own first character - "(kicad_pcb"
        // for the s-expression board and schematic, "{" for the JSON project file. A BOM or leading
        // whitespace from another editor is fine; a renamed program, an image, or the wrong KiCad
        // kind is not.
        // ###########################################################################################
        [Theory]
        [InlineData("board.kicad_pcb", "(kicad_pcb (version 20240108))", true)]
        [InlineData("board.kicad_sch", "(kicad_sch (version 20240108))", true)]
        [InlineData("board.kicad_pro", "{ \"board\": {} }", true)]
        [InlineData("board.kicad_pcb", "{ \"board\": {} }", false)]
        [InlineData("board.kicad_pro", "(kicad_pcb)", false)]
        [InlineData("board.kicad_sch", "just some text", false)]
        public void A_KiCad_file_opens_with_its_formats_own_character(string path, string content, bool accepted)
        {
            Assert.Equal(accepted, SubmissionContentRulesTests.Matches(path, Encoding.UTF8.GetBytes(content)));
        }

        [Fact]
        public void A_KiCad_file_may_open_with_a_BOM_and_whitespace_but_never_carry_a_NUL()
        {
            byte[] withBom = [0xEF, 0xBB, 0xBF, (byte)' ', (byte)'\r', (byte)'\n', .. Encoding.UTF8.GetBytes("(kicad_sch)")];
            Assert.True(SubmissionContentRulesTests.Matches("a.kicad_sch", withBom));

            byte[] binary = [.. Encoding.UTF8.GetBytes("(kicad_pcb"), 0, 0x4D, 0x5A];
            Assert.False(SubmissionContentRulesTests.Matches("a.kicad_pcb", binary));
        }

        // Images need only their signature; text needs the whole sample. Reading 64 KB of every one
        // of a board's thousand images would be waste, reading 32 bytes of a text file a blind spot.
        [Fact]
        public void The_bytes_needed_fit_the_type()
        {
            Assert.True(SubmissionContentRules.BytesNeeded("a.png") < 64);
            Assert.Equal(1024, SubmissionContentRules.BytesNeeded("a.pdf"));
            Assert.Equal(SubmissionContentRules.SampleBytes, SubmissionContentRules.BytesNeeded("a.txt"));
            Assert.Equal(SubmissionContentRules.SampleBytes, SubmissionContentRules.BytesNeeded("a.unknown"));
        }
    }
}
