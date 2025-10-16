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

namespace ReasonableRTF.Enums
{
    internal enum DestinationType : byte
    {
        FieldInstruction,
        /// <summary>
        /// This is for \csN, \dsN, and \tsN.
        /// <para/>
        /// These are weird hybrids that can either be written as destinations (eg. "\*\cs15") or not (eg. "\cs15").
        /// <para/>
        /// The spec explains:<br/>
        /// "\csN:<br/>
        /// Designates character style with a style handle N. Like \sN, \csN is not a destination control word.<br/>
        /// However, it is important to treat it like one inside the style sheet; that is, \csN must be prefixed<br/>
        /// with \* and must appear as the first item inside a group. Doing so ensures that readers that do not<br/>
        /// understand character styles will skip the character style information correctly. When used in body<br/>
        /// text to indicate that a character style was applied, do not include the \* prefix."
        /// <para/>
        /// Despite the fact that the \* prefixed versions are only supposed to appear at the start of a group,<br/>
        /// there's one readme (Thief Trinity) where \*\csN is written in the middle of a group. If we treated<br/>
        /// the \* prefixed version as being a skip-group trigger, then we would skip the rest of the group if it<br/>
        /// was in the middle of one, missing whatever text was after it.
        /// <para/>
        /// However, we actually don't have to treat any version of the word as a skip-group trigger, because the<br/>
        /// only time we want that is when they're in the \stylesheet group, which is already being skipped. So<br/>
        /// ignoring the word is a no-op in \stylesheet, and also a no-op in a regular group, which is what we want<br/>
        /// in both cases.
        /// </summary>
        CanBeDestOrNotDest,
        Skip,
        SkippableHex,
    }
}
