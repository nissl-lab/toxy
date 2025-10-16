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

namespace ReasonableRTF.Models
{
    // TODO: This is janky, maybe just let people use a custom symbol list? On the other hand this is way more convenient...
    public sealed class RtfToTextConverterOptions
    {
        /// <summary>
        /// Gets or sets whether to swap the uppercase and lowercase Greek phi characters in the Symbol font to Unicode
        /// translation table.
        /// <para/>
        /// The Windows Symbol font has these two characters swapped from their nominal positions.
        /// You can disable this by setting this property to <see langword="false"/>.
        /// <para/>
        /// The default value is <see langword="true"/>.
        /// </summary>
        public bool SwapUppercaseAndLowercasePhiSymbols { get; set; } = true;

        /// <summary>
        /// Gets or sets the character at index 0xA0 (160) in the Symbol font to Unicode translation table.
        /// This character is nominally the Euro sign, but in older versions of the Symbol font it may have been a
        /// numeric space or undefined.
        /// <para/>
        /// The default value is <see cref="SymbolFontA0Char.EuroSign"/>.
        /// </summary>
        public SymbolFontA0Char SymbolFontA0Char { get; set; } = SymbolFontA0Char.EuroSign;

        /// <summary>
        /// Gets or sets the linebreak style (CRLF or LF) for the converted plain text.
        /// <para/>
        /// The default value is <see cref="LineBreakStyle.CRLF"/>.
        /// </summary>
        public LineBreakStyle LineBreakStyle { get; set; } = LineBreakStyle.CRLF;

        /// <summary>
        /// Gets or sets whether to convert text that is marked as hidden. If <see langword="true"/>, this text will
        /// appear in the plain text output; otherwise it will not.
        /// <para/>
        /// The default value is <see langword="false"/>.
        /// </summary>
        public bool ConvertHiddenText { get; set; }

        internal void CopyTo(RtfToTextConverterOptions dest)
        {
            dest.SwapUppercaseAndLowercasePhiSymbols = SwapUppercaseAndLowercasePhiSymbols;
            dest.SymbolFontA0Char = SymbolFontA0Char;
            dest.LineBreakStyle = LineBreakStyle;
            dest.ConvertHiddenText = ConvertHiddenText;
        }
    }
}
