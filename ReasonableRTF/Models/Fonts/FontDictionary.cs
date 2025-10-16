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
using ReasonableRTF.Models.DataTypes;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace ReasonableRTF.Models.Fonts
{
    internal sealed class FontDictionary
    {
        private int _capacity;
        private readonly Dictionary<int, FontEntry> _dict;

        private readonly ListFast<FontEntry> _fontEntryPool;
        private int _fontEntryPoolVirtualCount;

        /*
        \fN params are normally in the signed int16 range, but the Windows RichEdit control supports them in the
        -30064771071 - 30064771070 (-0x6ffffffff - 0x6fffffffe) range (yes, bizarre numbers, but I tested and
        there they are).
        */

        internal FontEntry? Top;

        internal FontDictionary(int capacity)
        {
            _capacity = capacity;
            _fontEntryPool = new ListFast<FontEntry>(capacity);
            _dict = new Dictionary<int, FontEntry>(_capacity);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Add(int key)
        {
            FontEntry fontEntry;
            if (_fontEntryPoolVirtualCount > 0)
            {
                --_fontEntryPoolVirtualCount;
                fontEntry = _fontEntryPool[_fontEntryPoolVirtualCount];
                fontEntry.Reset();
            }
            else
            {
                fontEntry = new FontEntry();
                _fontEntryPool.Add(fontEntry);
            }

            Top = fontEntry;
            _dict[key] = fontEntry;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Clear()
        {
            Top = null;
            _dict.Clear();
            _fontEntryPoolVirtualCount = _fontEntryPool.Count;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void ClearFull(int capacity)
        {
            _capacity = capacity;
            _fontEntryPool.HardReset(capacity);
            _dict.Reset(capacity);
            Clear();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool TryGetValue(int key, [NotNullWhen(true)] out FontEntry? value)
        {
            return _dict.TryGetValue(key, out value);
        }
    }
}
