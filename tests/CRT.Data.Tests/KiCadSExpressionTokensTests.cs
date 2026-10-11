using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers KiCadSExpressionTokens - the one tokenizer CRT's KiCad loader and the server's KiCad
    // sheet rule both read with (security review, 2026-10-11).
    //
    // Characterisation tests: the tokenizer was moved here from the loader, its reading unchanged,
    // and the server's rule is only sound while it reads EXACTLY what the loader reads - quirks
    // included. It reads a TextReader a buffer at a time (code review, 2026-10-11), so every token
    // kind is also read across a buffer's edge.
    // ###########################################################################################
    public sealed class KiCadSExpressionTokensTests
    {
        // The reader's buffer, in characters - a token that starts this far in crosses its edge.
        private const int BufferSize = 16 * 1024;

        private static List<string> Read(string content, int maximumTokenLength = int.MaxValue)
        {
            var tokens = new KiCadSExpressionTokens(new StringReader(content), maximumTokenLength);
            var read = new List<string>();

            while (tokens.TryRead(out string? token))
                read.Add(token);

            return read;
        }

        private static string DataRoot()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "Assets", "Data");
        }

        // ###########################################################################################
        // The loader's tokenizer exactly as it was before it moved to CRT.Data (git: the private
        // SExpressionParser.Tokenize in KiCadRawProjectLoader), kept here as the reference the
        // reader must match. Never "fix" it - it is what CRTs already installed read.
        // ###########################################################################################
        private static List<string> ReferenceTokenize(string input)
        {
            var tokens = new List<string>();
            int i = 0;

            while (i < input.Length)
            {
                char ch = input[i];

                if (char.IsWhiteSpace(ch))
                {
                    i++;
                    continue;
                }

                if (ch == ';')
                {
                    while (i < input.Length && input[i] != '\n')
                    {
                        i++;
                    }

                    continue;
                }

                if (ch == '(' || ch == ')')
                {
                    tokens.Add(ch.ToString());
                    i++;
                    continue;
                }

                if (ch == '"')
                {
                    i++;
                    var value = new System.Text.StringBuilder();

                    while (i < input.Length)
                    {
                        char c = input[i];

                        if (c == '\\' && i + 1 < input.Length)
                        {
                            value.Append(input[i + 1]);
                            i += 2;
                            continue;
                        }

                        if (c == '"')
                        {
                            i++;
                            break;
                        }

                        value.Append(c);
                        i++;
                    }

                    tokens.Add(value.ToString());
                    continue;
                }

                int start = i;
                while (i < input.Length &&
                       !char.IsWhiteSpace(input[i]) &&
                       input[i] != '(' &&
                       input[i] != ')')
                {
                    i++;
                }

                tokens.Add(input[start..i]);
            }

            return tokens;
        }

        [Fact]
        public void Parentheses_atoms_and_quoted_strings_are_separate_tokens()
        {
            Assert.Equal(
                ["(", "property", "Sheetfile", "pages/cpu.kicad_sch", ")"],
                KiCadSExpressionTokens.Tokenize("(property \"Sheetfile\" pages/cpu.kicad_sch)"));
        }

        [Fact]
        public void A_quoted_string_keeps_its_spaces_and_loses_its_quotes()
        {
            Assert.Equal(["Power supply"], KiCadSExpressionTokens.Tokenize("\"Power supply\""));
        }

        // "\x" reads as x - so "\\\\server" in a file is the two backslashes of a network path.
        [Fact]
        public void A_backslash_in_a_quoted_string_escapes_the_next_character()
        {
            Assert.Equal(
                ["Net \"quoted\"", "\\\\server\\share"],
                KiCadSExpressionTokens.Tokenize("\"Net \\\"quoted\\\"\" \"\\\\\\\\server\\\\share\""));
        }

        // A backslash with nothing after it is kept, as the loader always kept it.
        [Fact]
        public void A_backslash_that_ends_the_file_is_kept()
        {
            Assert.Equal(["abc\\"], KiCadSExpressionTokens.Tokenize("\"abc\\"));
        }

        [Fact]
        public void A_semicolon_comment_runs_to_the_end_of_its_line()
        {
            Assert.Equal(
                ["(", "net", "1", ")", "(", "net", "2", ")"],
                KiCadSExpressionTokens.Tokenize("; header\n(net 1) ; trailing\n(net 2)"));
        }

        // Only a token's first character decides what it is: a quote or semicolon inside an atom is
        // part of it.
        [Fact]
        public void A_quote_or_semicolon_inside_an_atom_is_part_of_it()
        {
            Assert.Equal(["(", "a\"b", "c;d", ")"], KiCadSExpressionTokens.Tokenize("(a\"b c;d)"));
        }

        // The loader's quirk, kept on purpose: a quoted "(" is the same token as a real one.
        [Fact]
        public void A_quoted_parenthesis_is_the_same_token_as_a_real_one()
        {
            Assert.Equal(["(", "(", ")"], KiCadSExpressionTokens.Tokenize("\"(\" ( )"));
        }

        [Fact]
        public void An_unterminated_string_ends_at_the_end_of_the_file()
        {
            Assert.Equal(["(", "abc"], KiCadSExpressionTokens.Tokenize("( \"abc"));
        }

        [Fact]
        public void Empty_and_blank_text_have_no_tokens()
        {
            Assert.Empty(KiCadSExpressionTokens.Tokenize(string.Empty));
            Assert.Empty(KiCadSExpressionTokens.Tokenize(" \r\n\t "));
        }

        // A parse of a board's PCB meets millions of parentheses (code review, 2026-10-11).
        [Fact]
        public void Every_parenthesis_is_the_same_string()
        {
            List<string> tokens = KiCadSExpressionTokens.Tokenize("((a) (b))");

            Assert.All(tokens.Where(token => token == "("), token => Assert.Same(KiCadSExpressionTokens.Open, token));
            Assert.All(tokens.Where(token => token == ")"), token => Assert.Same(KiCadSExpressionTokens.Close, token));
        }

        // ###########################################################################################
        // *** A TOKEN ACROSS THE BUFFER'S EDGE READS AS THE LOADER READ IT (code review,
        // 2026-10-11). *** Each kind of token, a comment and an escape are placed to straddle the
        // edge of the first buffer and of the next, split at every point, and read against the
        // loader's own tokenizer as it was (ReferenceTokenize). Not theory data: a line break in a
        // test's name trips up the test runner.
        // ###########################################################################################
        [Fact]
        public void A_token_across_the_buffers_edge_reads_as_the_loader_read_it()
        {
            string[] texts =
            [
                "(an_atom_long_enough \"a quoted \\\" string\" ; a comment\n tail)",
                "\"\\\\\\\\server\\\\share\\\\x.kicad_sch\"",
                "atom\"with;inside",
                "\"unterminated \\"
            ];

            foreach (string text in texts)
            {
                foreach (int edge in new[] { KiCadSExpressionTokensTests.BufferSize, 2 * KiCadSExpressionTokensTests.BufferSize })
                {
                    for (int start = edge - text.Length - 1; start <= edge + 1; start++)
                    {
                        string content = "x" + new string(' ', start - 1) + text;

                        Assert.Equal(KiCadSExpressionTokensTests.ReferenceTokenize(content), KiCadSExpressionTokensTests.Read(content));
                    }
                }
            }
        }

        // ###########################################################################################
        // *** EVERY SHIPPED KiCad FILE, READ AS THE SERVER READS IT, GIVES THE TOKENS THE LOADER GOT
        // FROM IT (code review, 2026-10-11). *** The server streams a blob through a StreamReader;
        // the loader tokenized the file's whole text. The sheet rule is only sound while the two
        // agree on every real file.
        //
        // Files over 8 MB are left out: the one shipped (a 44 MB PCB, mostly zone outlines) holds no
        // kind of token the others lack, and alone took most of the test's three seconds.
        // ###########################################################################################
        [Fact]
        public void Every_shipped_KiCad_file_streamed_gives_the_tokens_the_loader_read()
        {
            string dataRoot = KiCadSExpressionTokensTests.DataRoot();
            int files = 0;

            foreach (string path in Directory.EnumerateFiles(dataRoot, "*.kicad_*", SearchOption.AllDirectories)
                         .Where(path => path.EndsWith(".kicad_sch", StringComparison.OrdinalIgnoreCase) ||
                                        path.EndsWith(".kicad_pcb", StringComparison.OrdinalIgnoreCase))
                         .Where(path => new FileInfo(path).Length <= 8 * 1024 * 1024))
            {
                List<string> loaderRead = KiCadSExpressionTokensTests.ReferenceTokenize(File.ReadAllText(path));

                using var stream = new StreamReader(path, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                var tokens = new KiCadSExpressionTokens(stream);
                int index = 0;

                while (tokens.TryRead(out string? token))
                {
                    Assert.True(index < loaderRead.Count, $"{path}: more tokens than the loader read.");
                    Assert.True(loaderRead[index] == token, $"{path}: token {index} is [{token}], the loader read [{loaderRead[index]}].");
                    index++;
                }

                Assert.Equal(loaderRead.Count, index);
                files++;
            }

            Assert.True(files > 0, "No shipped KiCad file was found.");
        }

        // An atom or string longer than the buffer itself.
        [Fact]
        public void A_token_longer_than_the_buffer_reads_whole()
        {
            string atom = new('a', 3 * KiCadSExpressionTokensTests.BufferSize + 7);
            string quoted = new('q', 2 * KiCadSExpressionTokensTests.BufferSize + 3);

            Assert.Equal(["(", atom, quoted, ")"], KiCadSExpressionTokens.Tokenize($"({atom} \"{quoted}\")"));
        }

        // ###########################################################################################
        // *** A READER WITH A MAXIMUM HOLDS NO TOKEN WHOLE (code review, 2026-10-11). *** The server
        // checks a schematic of up to the blob store's limit; a token longer than the maximum comes
        // back one character past it - known to be too long - and the tokens after it are unharmed.
        // ###########################################################################################
        [Fact]
        public void A_token_longer_than_the_maximum_comes_back_one_character_past_it()
        {
            string atom = new('a', 3 * KiCadSExpressionTokensTests.BufferSize);
            string quoted = new('q', 50);

            Assert.Equal(
                ["(", new string('a', 11), new string('q', 11), "short", "exactly10!", ")"],
                KiCadSExpressionTokensTests.Read($"({atom} \"{quoted}\" short exactly10!)", maximumTokenLength: 10));
        }

        [Fact]
        public void A_maximum_below_one_character_is_refused()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new KiCadSExpressionTokens(new StringReader("x"), 0));
        }

        // Tokenize reads the same tokens as a reader of the same text - the loader's list is the
        // server's stream.
        [Fact]
        public void Tokenize_gives_the_same_tokens_as_a_reader()
        {
            const string Content = "(kicad_sch (version 20231120) ; note\n (sheet (property \"Sheetname\" \"CPU\")))";

            Assert.Equal(KiCadSExpressionTokensTests.Read(Content), KiCadSExpressionTokens.Tokenize(Content));
        }
    }
}
