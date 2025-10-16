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

using ReasonableRTF.Models.DataTypes;
using System.Runtime.CompilerServices;

namespace ReasonableRTF.Helper
{
    internal static class UtilHelper
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool IsAsciiHex(this byte b) => char.IsAsciiHexDigit((char)b);

        /// <summary>
        /// Returns an array of type <typeparamref name="T"/> with all elements initialized to <paramref name="value"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="length"></param>
        /// <param name="value">The value to initialize all elements with.</param>
        internal static T[] InitializedArray<T>(int length, T value) where T : new()
        {
            T[] ret = new T[length];
            for (int i = 0; i < length; i++)
            {
                ret[i] = value;
            }
            return ret;
        }

        /// <summary>
        /// Clears the dictionary and sets its internal storage to zero-length.
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="dictionary"></param>
        /// <param name="capacity"></param>
        internal static void Reset<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, int capacity) where TKey : notnull
        {
            dictionary.Clear();
            dictionary.TrimExcess(capacity);
        }

        /// <summary>
        /// Copy of .NET 7 version (fewer branches than Framework) but with a fast null return on fail instead of the infernal exception-throwing.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static ListFast<char>? ConvertFromUtf32(uint utf32u, ListFast<char> charBuffer)
        {
            if ((utf32u - 0x110000u ^ 0xD800u) < 0xFFEF0800u)
            {
                return null;
            }

            if (utf32u <= char.MaxValue)
            {
                charBuffer.ItemsArray[0] = (char)utf32u;
                charBuffer.Count = 1;
                return charBuffer;
            }

            charBuffer.ItemsArray[0] = (char)(utf32u + (0xD800u - 0x40u << 10) >> 10);
            charBuffer.ItemsArray[1] = (char)((utf32u & 0x3FFu) + 0xDC00u);
            charBuffer.Count = 2;

            return charBuffer;
        }
    }
}
