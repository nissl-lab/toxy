/*
 * MIT License
 * 
 * Copyright (c) 2024 Brian Tobin
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
*/

using ReasonableRTF.Enums;
using System.Runtime.CompilerServices;

namespace ReasonableRTF.Models.Symbols
{
    internal sealed class SymbolDict
    {
        /* ANSI-C code produced by gperf version 3.1 */
        /* Command-line: 'C:\\gperf\\tools\\gperf.exe' --output-file='C:\\_al_rtf_table_gen\\gperfOutputFile.txt' -t 'C:\\_al_rtf_table_gen\\gperfFormatFile.txt'  */
        /* Computed positions: -k'1-3,$' */

        //private const int TOTAL_KEYWORDS = 83;
        //private const int MIN_WORD_LENGTH = 1;
        private const int MAX_WORD_LENGTH = 18;
        //private const int MIN_HASH_VALUE = 11;
        private const int MAX_HASH_VALUE = 283;
        /* maximum key range = 273, duplicates = 0 */

        private readonly ushort[] asso_values =
        [
            284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 10, 45, 10,
        75, 10, 5, 95, 40, 5, 90, 0, 0, 35,
        15, 0, 25, 15, 50, 15, 0, 70, 10, 100,
        32, 0, 0, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284, 284, 284, 284, 284,
        284, 284, 284, 284, 284, 284,
    ];

        /*
        For "cs", "ds", "ts"
        Hack to make sure we extract the \fldrslt text from Thief Trinity in that one place.

        For "listtext", "pntext"
        TODO(listtext/pntext): Temporarily disabled with a hack, but decide what we want to do here

        For "v"
        \v to make all plain text hidden (not output to the conversion stream), \v0 to make it shown again

        For "ansi"
        The spec calls this "ANSI (the default)" but says nothing about what codepage that actually means.
        "ANSI" is often misused to mean one of the Windows codepages, so I'll assume it's Windows-1252.

        For "mac"
        The spec calls this "Apple Macintosh" but again says nothing about what codepage that is. I'll
        assume 10000 ("Mac Roman")

        For "fldinst"
        We need to do stuff with this (SYMBOL instruction)

        NOTE: This is generated. Values can be modified, but not keys (keys are the first string params).
        Also no reordering. Adding, removing, reordering, or modifying keys requires generating a new version.
        See RTF_SymbolListGenSource.cs for how to generate a new version (it also contains the original
        Symbol list which must be used as the source to generate this one).
        */
        private readonly Symbol?[] _symbolTable =
        [
            null, null, null, null, null, null, null, null, null,
        null, null,
// Entry 7
        new Symbol("f", 0, false, KeywordType.Property, (ushort)Property.FontNum),
// Entry 47
        new Symbol("footerl", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null, null, null,
// Entry 46
        new Symbol("footerf", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 41
        new Symbol("colortbl", 0, false, KeywordType.Special, (ushort)SpecialType.ColorTable),
        null,
// Entry 67
        new Symbol("title", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 13
        new Symbol("v", 1, false, KeywordType.Property, (ushort)Property.Hidden),
// Entry 66
        new Symbol("tc", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 49
        new Symbol("footnote", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 80
        new Symbol("cell", 0, false, KeywordType.Character, '\t'),
// Entry 64
        new Symbol("stylesheet", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 6
        new Symbol("fonttbl", 0, false, KeywordType.Special, (ushort)SpecialType.FontTable),
// Entry 37
        new Symbol("listtext", 0, false, KeywordType.Destination, 255),
// Entry 57
        new Symbol("info", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null,
// Entry 36
        new Symbol("ts", 0, false, KeywordType.Destination, (ushort)DestinationType.CanBeDestOrNotDest),
// Entry 58
        new Symbol("keywords", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 15
        new Symbol("line", 0, false, KeywordType.Character, '\n'),
        null, null,
// Entry 52
        new Symbol("ftnsepc", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 16
        new Symbol("sect", 0, false, KeywordType.Character, '\n'),
// Entry 50
        new Symbol("ftncn", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 34
        new Symbol("cs", 0, false, KeywordType.Destination, (ushort)DestinationType.CanBeDestOrNotDest),
        null,
// Entry 70
        new Symbol("pict", 1, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
        null,
// Entry 38
        new Symbol("pntext", 0, false, KeywordType.Destination, 255),
// Entry 1
        new Symbol("pc", 437, true, KeywordType.Special, (ushort)SpecialType.HeaderCodePage),
// Entry 82
        new Symbol("nestcell", 0, false, KeywordType.Character, '\t'),
// Entry 0
        new Symbol("ansi", 1252, true, KeywordType.Special, (ushort)SpecialType.HeaderCodePage),
        null,
// Entry 51
        new Symbol("ftnsep", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 42
        new Symbol("comment", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 69
        new Symbol("xe", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 68
        new Symbol("txe", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 20
        new Symbol("enspace", 0, false, KeywordType.Character, '\x2002'),
// Entry 3
        new Symbol("pca", 850, true, KeywordType.Special, (ushort)SpecialType.HeaderCodePage),
        null, null,
// Entry 45
        new Symbol("footer", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 48
        new Symbol("footerr", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 8
        new Symbol("fcharset", -1, false, KeywordType.Special, (ushort)SpecialType.Charset),
        null, null,
// Entry 78
        new Symbol("panose", 20, true, KeywordType.Special, (ushort)SpecialType.SkipNumberOfBytes),
// Entry 55
        new Symbol("headerl", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 2
        new Symbol("mac", 10000, true, KeywordType.Special, (ushort)SpecialType.HeaderCodePage),
// Entry 71
        new Symbol("themedata", 0, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
        null, null,
// Entry 54
        new Symbol("headerf", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null, null, null,
// Entry 19
        new Symbol("emspace", 0, false, KeywordType.Character, '\x2003'),
        null, null, null, null,
// Entry 21
        new Symbol("qmspace", 0, false, KeywordType.Character, '\x2005'),
// Entry 32
        new Symbol("bin", 0, false, KeywordType.Special, (ushort)SpecialType.SkipNumberOfBytes),
        null, null, null,
// Entry 33
        new Symbol("fldinst", 0, false, KeywordType.Destination, (ushort)DestinationType.FieldInstruction),
        null, null, null, null,
// Entry 11
        new Symbol("uc", 1, false, KeywordType.Property, (ushort)Property.UnicodeCharSkipCount),
// Entry 59
        new Symbol("operator", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null, null,
// Entry 61
        new Symbol("private", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 5
        new Symbol("deff", 0, false, KeywordType.Special, (ushort)SpecialType.DefaultFont),
        null,
// Entry 24
        new Symbol("lquote", 0, false, KeywordType.Character, '\x2018'),
// Entry 73
        new Symbol("passwordhash", 0, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
// Entry 17
        new Symbol("tab", 0, false, KeywordType.Character, '\t'),
// Entry 74
        new Symbol("datastore", 0, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
// Entry 63
        new Symbol("rxe", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null,
// Entry 35
        new Symbol("ds", 0, false, KeywordType.Destination, (ushort)DestinationType.CanBeDestOrNotDest),
        null, null, null,
// Entry 62
        new Symbol("revtim", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 43
        new Symbol("creatim", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null, null,
// Entry 53
        new Symbol("header", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 56
        new Symbol("headerr", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null,
// Entry 29
        new Symbol("zwnbo", 0, false, KeywordType.Character, '\xFEFF'),
// Entry 18
        new Symbol("bullet", 0, false, KeywordType.Character, '\x2022'),
// Entry 60
        new Symbol("printim", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 72
        new Symbol("colorschememapping", 0, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
// Entry 10
        new Symbol("lang", 0, false, KeywordType.Property, (ushort)Property.Lang),
        null, null,
// Entry 44
        new Symbol("doccomm", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null, null, null,
// Entry 77
        new Symbol("blipuid", 32, true, KeywordType.Special, (ushort)SpecialType.SkipNumberOfBytes),
        null, null, null,
// Entry 39
        new Symbol("author", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 65
        new Symbol("subject", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
// Entry 14
        new Symbol("par", 0, false, KeywordType.Character, '\n'),
// Entry 26
        new Symbol("ldblquote", 0, false, KeywordType.Character, '\x201C'),
        null,
// Entry 12
        new Symbol("u", 0, false, KeywordType.Special, (ushort)SpecialType.UnicodeChar),
// Entry 4
        new Symbol("ansicpg", 1252, false, KeywordType.Special, (ushort)SpecialType.HeaderCodePage),
        null, null, null,
// Entry 23
        new Symbol("endash", 0, false, KeywordType.Character, '\x2013'),
// Entry 81
        new Symbol("nestrow", 0, false, KeywordType.Special, (ushort)SpecialType.CellRowEnd),
        null,
// Entry 28
        new Symbol("zwbo", 0, false, KeywordType.Character, '\x200B'),
        null,
// Entry 25
        new Symbol("rquote", 0, false, KeywordType.Character, '\x2019'),
// Entry 76
        new Symbol("objdata", 1, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
        null, null, null, null, null, null, null, null, null,
        null, null, null, null,
// Entry 22
        new Symbol("emdash", 0, false, KeywordType.Character, '\x2014'),
        null, null,
// Entry 75
        new Symbol("datafield", 0, false, KeywordType.Destination, (ushort)DestinationType.SkippableHex),
        null, null, null, null, null, null, null, null, null,
        null, null,
// Entry 40
        new Symbol("buptim", 0, false, KeywordType.Destination, (ushort)DestinationType.Skip),
        null, null, null, null, null, null, null,
// Entry 27
        new Symbol("rdblquote", 0, false, KeywordType.Character, '\x201D'),
        null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null, null, null, null,
        null,
// Entry 31
        new Symbol("zwnj", 0, false, KeywordType.Character, '\x200C'),
        null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null, null, null, null,
// Entry 9
        new Symbol("cpg", -1, false, KeywordType.Special, (ushort)SpecialType.CodePage),
        null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null,
// Entry 79
        new Symbol("row", 0, false, KeywordType.Special, (ushort)SpecialType.CellRowEnd),
        null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null, null, null, null,
        null, null,
// Entry 30
        new Symbol("zwj", 0, false, KeywordType.Character, '\x200D'),
    ];

        private static Symbol?[] InitControlSymbolArray()
        {
            Symbol?[] ret = new Symbol?[256];
            ret['\''] = new Symbol("'", 0, false, KeywordType.Special, (int)SpecialType.HexEncodedChar);
            /*
            @RTF(KeywordType.Character and symbol fonts):
            \, {, and } are the only KeywordType.Character chars that can be in a symbol font. Everything else is
            either below 0x20 or more than one byte, which in either case means they can't be symbol font chars.
            ~ is nominally a non-breaking space, and in RichEdit is displayed as such (or at least whitespace of
            some kind), but in LibreOffice is displayed as a square dot when set to Wingdings (as expected).
            Since RichEdit doesn't treat it as a symbol font character we should in theory match its behavior,
            but we convert it to an ASCII space anyway so the whole thing is moot currently. But just in case we
            decide to change it, there's the info.

            We could maybe figure out a way to not have to do the symbol font check/conversion in the common case
            where we don't need to, is the point of this whole soliloquy.

            TODO: Handle bulleted/numbered lists properly
            TODO: "ansi" keyword should be system default ANSI codepage maybe? 1252 currently
            TODO: Remove all assumptions about Windows-1252
            NOTE(Footnotes):
            RichTextBox doesn't convert them at all.
            LibreOffice adds citation numbers but doesn't add the footnotes themselves.
            */
            ret['\\'] = new Symbol("\\", 0, false, KeywordType.Character, '\\');
            ret['{'] = new Symbol("{", 0, false, KeywordType.Character, '{');
            ret['}'] = new Symbol("}", 0, false, KeywordType.Character, '}');

            // Non-breaking space (0xA0)
            ret['~'] = new Symbol("~", 0, false, KeywordType.Character, '\xA0');

            // Non-breaking hyphen (0x2011)
            ret['_'] = new Symbol("_", 0, false, KeywordType.Character, '\x2011');

            // Soft hyphen (Spec calls this "Optional hyphen")
            ret['-'] = new Symbol("-", 0, false, KeywordType.Character, '\xAD');

            // There's also \: which "specifies a subentry in an index entry" (it's not clear even from the spec what
            // exactly an "index entry" is).

            /*
            Spec:
            "A carriage return (character value 13) or line feed (character value 10) is treated as a \par
            control if the character is preceded by a backslash. You must include the backslash; otherwise,
            RTF ignores the control word."
            */
            ret['\r'] = new Symbol("\r", 0, false, KeywordType.Character, '\n');
            ret['\n'] = new Symbol("\n", 0, false, KeywordType.Character, '\n');
            return ret;
        }

        private readonly Symbol?[] _controlSymbols = InitControlSymbolArray();

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal Symbol? LookUpControlSymbol(char ch) => _controlSymbols[ch];

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal Symbol? LookUpControlWord(char[] keyword, int len)
        {
            // Min word length is 1, and we're guaranteed to always be at least 1, so no need to check for >= min
            if (len <= MAX_WORD_LENGTH)
            {
                int key = len;

                // Original C code does a stupid thing where it puts default at the top and falls through and junk,
                // but we can't do that in C#, so have something clearer/clunkier
                switch (len)
                {
                    // Most common case first - we get a measurable speedup from this
                    case > 2:
                        key += asso_values[keyword[2]];
                        key += asso_values[keyword[1]];
                        key += asso_values[keyword[0]];
                        break;
                    case 1:
                        key += asso_values[keyword[0]];
                        break;
                    case 2:
                        key += asso_values[keyword[1]];
                        key += asso_values[keyword[0]];
                        break;
                }
                key += asso_values[keyword[len - 1]];

                if (key <= MAX_HASH_VALUE)
                {
                    Symbol? symbol = _symbolTable[key];
                    if (symbol == null)
                    {
                        return null;
                    }

                    string seq2 = symbol.Keyword;
                    if (len != seq2.Length)
                    {
                        return null;
                    }

                    for (int ci = 0; ci < len; ci++)
                    {
                        if (keyword[ci] != seq2[ci])
                        {
                            return null;
                        }
                    }

                    return symbol;
                }
            }

            return null;
        }
    }
}
