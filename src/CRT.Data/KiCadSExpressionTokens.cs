using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The tokens of a KiCad s-expression file (.kicad_sch, .kicad_pcb) - the ONE tokenizer CRT
    // reads KiCad files with (security review, 2026-10-11).
    //
    // CRT's KiCad loader (KiCadRawProjectLoader) builds its trees from these tokens, and the
    // server's rule for a schematic's child sheets (KiCadSheetReferences) reads the same ones, so
    // the two cannot disagree about what a file says: a sheet path the server checked is the path
    // CRT would follow. Moved here from the loader, its reading unchanged.
    //
    // A token is a parenthesis, a bare atom, or a quoted string with its quotes removed and its
    // backslash escapes resolved ("\x" reads as x). ';' starts a comment running to the end of
    // the line.
    //
    // *** A QUOTED "(" OR ")" IS THE SAME TOKEN AS A REAL ONE. *** That is how the loader has
    // always read it, and it is kept rather than fixed: the rule above must read exactly what the
    // loader reads.
    //
    // READ FROM A TextReader, a buffer at a time (code review, 2026-10-11): the server checks a
    // submitted schematic of up to the blob store's size limit, and holding all of it as one string
    // took twice its size in memory. A reader that sets `maximumTokenLength` holds no more than that
    // of any one token either. The loader reads its strings through here too (Tokenize), so there is
    // still one tokenizer. Both parentheses are always the same two strings, never a new one each.
    // ###########################################################################################
    public sealed class KiCadSExpressionTokens
    {
        public const string Open = "(";

        public const string Close = ")";

        private const int BufferSize = 16 * 1024;

        private readonly TextReader thisText;

        // The most characters a token keeps - one past the caller's maximum, so a token that was
        // cut is still known to be too long (KiCadSheetReferences refuses one).
        private readonly int thisKept;

        private readonly char[] thisBuffer = new char[KiCadSExpressionTokens.BufferSize];

        // A quoted string, or an atom running past the end of the buffer, is gathered here.
        private readonly StringBuilder thisToken = new();

        private int thisPosition;
        private int thisLength;

        // ###########################################################################################
        // `maximumTokenLength`: a longer token is cut to one character more than this, and the rest
        // of it is read past without being kept. Every token is whole by default.
        // ###########################################################################################
        public KiCadSExpressionTokens(TextReader text, int maximumTokenLength = int.MaxValue)
        {
            ArgumentNullException.ThrowIfNull(text);
            ArgumentOutOfRangeException.ThrowIfLessThan(maximumTokenLength, 1);

            this.thisText = text;
            this.thisKept = maximumTokenLength == int.MaxValue ? int.MaxValue : maximumTokenLength + 1;
        }

        // ###########################################################################################
        // Every token of the file, in order - what the loader parses its trees from.
        //
        // The list starts at a size for a typical board: the shipped boards' files run 4 to 6
        // characters per token, apart from one PCB of mostly long zone outlines at 40.
        // ###########################################################################################
        public static List<string> Tokenize(string input)
        {
            ArgumentNullException.ThrowIfNull(input);

            var tokens = new List<string>(input.Length / 6);
            var reader = new KiCadSExpressionTokens(new StringReader(input));

            while (reader.TryRead(out string? token))
            {
                tokens.Add(token);
            }

            return tokens;
        }

        // ###########################################################################################
        // The next token, or false at the end of the text.
        // ###########################################################################################
        public bool TryRead([NotNullWhen(true)] out string? token)
        {
            while (this.TryPeek(out char ch))
            {
                if (char.IsWhiteSpace(ch))
                {
                    this.thisPosition++;
                    continue;
                }

                if (ch == ';')
                {
                    while (this.TryPeek(out char inComment) && inComment != '\n')
                    {
                        this.thisPosition++;
                    }

                    continue;
                }

                if (ch == '(' || ch == ')')
                {
                    this.thisPosition++;
                    token = ch == '(' ? KiCadSExpressionTokens.Open : KiCadSExpressionTokens.Close;
                    return true;
                }

                if (ch == '"')
                {
                    this.thisPosition++;
                    token = this.ReadQuoted();
                    return true;
                }

                token = this.ReadAtom();
                return true;
            }

            token = null;
            return false;
        }

        // ###########################################################################################
        // A quoted string, its opening quote already read. "\x" reads as x - but a backslash that
        // ends the file is kept as it is - and a string the file never closes ends with the file.
        // ###########################################################################################
        private string ReadQuoted()
        {
            this.thisToken.Clear();

            while (this.TryPeek(out char ch))
            {
                this.thisPosition++;

                if (ch == '"')
                {
                    break;
                }

                if (ch == '\\' && this.TryPeek(out char escaped))
                {
                    this.thisPosition++;
                    ch = escaped;
                }

                if (this.thisToken.Length < this.thisKept)
                {
                    this.thisToken.Append(ch);
                }
            }

            return this.thisToken.ToString();
        }

        // ###########################################################################################
        // A bare atom: everything up to whitespace or a parenthesis. A '"' or ';' inside one is
        // part of it - only the first character decides what a token is.
        // ###########################################################################################
        private string ReadAtom()
        {
            this.thisToken.Clear();

            while (true)
            {
                int start = this.thisPosition;

                while (this.thisPosition < this.thisLength && !KiCadSExpressionTokens.EndsAtom(this.thisBuffer[this.thisPosition]))
                {
                    this.thisPosition++;
                }

                bool ended = this.thisPosition < this.thisLength;
                int count = this.thisPosition - start;

                // The usual case: the whole atom inside the buffer, made into its string directly.
                if (ended && this.thisToken.Length == 0)
                {
                    return new string(this.thisBuffer, start, Math.Min(count, this.thisKept));
                }

                int room = this.thisKept - this.thisToken.Length;

                if (room > 0)
                {
                    this.thisToken.Append(this.thisBuffer, start, Math.Min(count, room));
                }

                if (ended || !this.Fill())
                {
                    return this.thisToken.ToString();
                }
            }
        }

        private static bool EndsAtom(char ch) => char.IsWhiteSpace(ch) || ch == '(' || ch == ')';

        // The character at the read position, refilling the buffer when it is used up.
        private bool TryPeek(out char ch)
        {
            if (this.thisPosition >= this.thisLength && !this.Fill())
            {
                ch = default;
                return false;
            }

            ch = this.thisBuffer[this.thisPosition];
            return true;
        }

        private bool Fill()
        {
            this.thisPosition = 0;
            this.thisLength = this.thisText.Read(this.thisBuffer, 0, this.thisBuffer.Length);

            return this.thisLength > 0;
        }
    }
}
