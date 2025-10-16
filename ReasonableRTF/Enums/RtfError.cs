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
    // TODO: Make all the names and descriptions clear for public API purposes.
    public enum RtfError : byte
    {
        /// <summary>
        /// No error.
        /// </summary>
        OK,
        /// <summary>
        /// The file did not have a valid rtf header.
        /// </summary>
        NotAnRtfFile,
        /// <summary>
        /// Unmatched '}'.
        /// </summary>
        StackUnderflow,
        /// <summary>
        /// Unmatched '{'.
        /// </summary>
        UnmatchedBrace,
        /// <summary>
        /// End of file was unexpectedly encountered while parsing.
        /// </summary>
        UnexpectedEndOfFile,
        /// <summary>
        /// A keyword longer than 32 characters was encountered.
        /// </summary>
        KeywordTooLong,
        /// <summary>
        /// A keyword parameter was outside the range of -2147483648 to 2147483647, or was longer than 10 characters.
        /// </summary>
        ParameterOutOfRange,
        /// <summary>
        /// The rtf was malformed in such a way that it might have been unsafe to continue parsing it (infinite loops, stack overflows, etc.)
        /// </summary>
        AbortedForSafety,
        /// <summary>
        /// An unexpected error occurred.
        /// </summary>
        UnexpectedError,
    }
}
