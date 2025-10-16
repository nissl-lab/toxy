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

namespace ReasonableRTF.Models.Symbols
{
    internal sealed class Symbol
    {
        internal readonly string Keyword;
        internal readonly int DefaultParam;
        internal readonly bool UseDefaultParam;
        internal readonly KeywordType KeywordType;
        /// <summary>
        /// Index into the property table, or a regular enum member, or a character literal, depending on <see cref="KeywordType"/>.
        /// </summary>
        internal readonly ushort Index;

        internal Symbol(string keyword, int defaultParam, bool useDefaultParam, KeywordType keywordType, ushort index)
        {
            Keyword = keyword;
            DefaultParam = defaultParam;
            UseDefaultParam = useDefaultParam;
            KeywordType = keywordType;
            Index = index;
        }
    }
}
