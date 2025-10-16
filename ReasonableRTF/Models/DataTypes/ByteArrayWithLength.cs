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

using ReasonableRTF.Helper;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ReasonableRTF.Models.DataTypes
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct ByteArrayWithLength
    {
        internal readonly byte[] Array;
        internal readonly int Length;

        public ByteArrayWithLength()
        {
            Array = System.Array.Empty<byte>();
            Length = 0;
        }

        internal ByteArrayWithLength(byte[] array)
        {
            Array = array;
            Length = array.Length;
        }

        internal ByteArrayWithLength(byte[] array, int length)
        {
            Array = array;
            Length = length;
        }

        // This MUST be a method (not a static field) to maintain performance!
        internal static ByteArrayWithLength Empty() => new();

        /// <summary>
        /// Manually bounds-checked past <see cref="T:Length"/>.
        /// If you don't need bounds-checking past <see cref="T:Length"/>, access <see cref="T:Array"/> directly.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        internal byte this[int index]
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get
            {
                // Very unfortunately, we have to manually bounds-check here, because our array could be longer
                // than Length (such as when it comes from a pool).
                if (index > Length - 1) ThrowHelper.IndexOutOfRange();
                return Array[index];
            }
        }
    }
}
