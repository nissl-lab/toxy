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
using ReasonableRTF.Attributes;
using System.Runtime.CompilerServices;

namespace ReasonableRTF.Models
{
    internal sealed class GroupStack
    {
        // The Max Length of an Array 
        // Copy of https://github.com/dotnet/runtime/blob/main/src/libraries/System.Private.CoreLib/src/System/Array.cs
        private const int MaxLength = 0X7FFFFFC7;
        internal const int PropertiesLen = 4;
        private const int DefaultCapacity = 100;
        private int Capacity;

        private bool[] _skipDestinations;
        private bool[] _inFontTables;
        internal byte[] _symbolFonts;
        internal int[][] Properties;

        internal int Count;

#nullable disable
        internal GroupStack()
        {
            Init();
        }
#nullable restore

        [MemberNotNull(
            nameof(_skipDestinations),
            nameof(_inFontTables),
            nameof(_symbolFonts),
            nameof(Properties))]
        private void Init()
        {
            Count = 0;
            Capacity = DefaultCapacity;

            _skipDestinations = new bool[Capacity];
            _inFontTables = new bool[Capacity];
            _symbolFonts = new byte[Capacity];
            Properties = new int[Capacity][];

            for (int i = 0; i < Capacity; i++)
            {
                Properties[i] = new int[PropertiesLen];
            }
        }

        private void Grow()
        {
            int oldMaxGroups = Capacity;

            int newCapacity = Capacity * 2;
            if ((uint)newCapacity > MaxLength) newCapacity = MaxLength;

            Capacity = newCapacity;
            Array.Resize(ref _skipDestinations, Capacity);
            Array.Resize(ref _inFontTables, Capacity);
            Array.Resize(ref _symbolFonts, Capacity);
            Array.Resize(ref Properties, Capacity);

            for (int i = oldMaxGroups; i < Capacity; i++)
            {
                Properties[i] = new int[PropertiesLen];
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void DeepCopyToNext()
        {
            // We don't really take a speed hit from this at all, but we support files with a stupid amount of
            // nested groups now.
            if (Count >= Capacity - 1)
            {
                Grow();
            }

            _skipDestinations[Count + 1] = _skipDestinations[Count];
            _inFontTables[Count + 1] = _inFontTables[Count];
            _symbolFonts[Count + 1] = _symbolFonts[Count];
            for (int i = 0; i < PropertiesLen; i++)
            {
                Properties[Count + 1][i] = Properties[Count][i];
            }
            ++Count;
        }

        #region Current group

        internal bool CurrentSkipDest
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get => _skipDestinations[Count];
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            set => _skipDestinations[Count] = value;
        }

        internal bool CurrentInFontTable
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get => _inFontTables[Count];
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            set => _inFontTables[Count] = value;
        }

        internal SymbolFont CurrentSymbolFont
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get => (SymbolFont)_symbolFonts[Count];
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            set => _symbolFonts[Count] = (byte)value;
        }

        internal int[] CurrentProperties
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get => Properties[Count];
        }

        // Current group always begins at group 0, so reset just that one
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void ResetFirst()
        {
            _skipDestinations[0] = false;
            _inFontTables[0] = false;
            _symbolFonts[0] = (int)SymbolFont.None;

            Properties[0][(int)Property.Hidden] = 0;
            Properties[0][(int)Property.UnicodeCharSkipCount] = 1;
            Properties[0][(int)Property.FontNum] = RtfToTextConverter.NoFontNumber;
            Properties[0][(int)Property.Lang] = -1;
        }

        #endregion

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void ClearFast() => Count = 0;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void ResetCapacityIfTooHigh()
        {
            if (Capacity > DefaultCapacity)
            {
                Init();
            }
        }
    }
}
