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

using System.Runtime.CompilerServices;

namespace ReasonableRTF.Models.DataTypes
{
    // How many times have you thought, "Gee, I wish I could just reach in and grab that backing array from
    // that List, instead of taking the senseless performance hit of having it copied to a newly allocated
    // array all the time in a ToArray() call"? Hooray!
    /// <summary>
    /// Because this list exposes its internal array and also doesn't clear said array on <see cref="ClearFast"/>,
    /// it must be used with care.
    /// <para>
    /// -Only use this with value types. Reference types will be left hanging around in the array.
    /// </para>
    /// <para>
    /// -The internal array is there so you can get at it without incurring an allocation+copy.
    ///  It can very easily become desynced with the <see cref="ListFast{T}"/> if you modify it.
    /// </para>
    /// <para>
    /// -Only use the internal array in conjunction with the <see cref="Count"/> property.
    ///  Using the <see cref="ItemsArray"/>.Length value will get the array's actual length, when what you
    ///  wanted was the list's "virtual" length. This is the same as a normal List except with a normal List
    ///  the array is private so you can't have that problem.
    /// </para>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal sealed class ListFast<T>
    {
        internal T[] ItemsArray;
        private int _itemsArrayLength;

        /// <summary>
        /// Properties are slow. You can set this from outside if you know what you're doing.
        /// </summary>
        internal int Count;

        /// <summary>
        /// No bounds checking, so use caution!
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        internal T this[int index]
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get => ItemsArray[index];
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            set => ItemsArray[index] = value;
        }

        internal ListFast(int capacity)
        {
            ItemsArray = new T[capacity];
            _itemsArrayLength = capacity;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Add(T item)
        {
            if (Count == _itemsArrayLength) EnsureCapacity(Count + 1);
            ItemsArray[Count++] = item;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void AddRange(ListFast<T> items, int count)
        {
            EnsureCapacity(Count + count);
            // We usually add small enough arrays that a loop is faster
            for (int i = 0; i < count; i++)
            {
                ItemsArray[Count + i] = items[i];
            }
            Count += count;
        }

        /*
        Honestly, for fixed-size value types, doing an Array.Clear() is completely unnecessary. For reference
        types, you definitely want to clear it to get rid of all the references, but for ints or chars etc.,
        all a clear does is set a bunch of fixed-width values to other fixed-width values. You don't save
        space and you don't get rid of loose references, all you do is waste an alarming amount of time. We
        drop fully 200ms from the Unicode parser just by using the fast clear!
        */
        /// <summary>
        /// Just sets <see cref="Count"/> to 0. Doesn't zero out the array or do anything else whatsoever.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void ClearFast() => Count = 0;

        internal int Capacity
        {
            get => _itemsArrayLength;
            set
            {
                if (value == _itemsArrayLength) return;
                if (value > 0)
                {
                    T[] objArray = new T[value];
                    if (Count > 0) Array.Copy(ItemsArray, 0, objArray, 0, Count);
                    ItemsArray = objArray;
                    _itemsArrayLength = value;
                    if (_itemsArrayLength < Count) Count = _itemsArrayLength;
                }
                else
                {
                    ItemsArray = Array.Empty<T>();
                    _itemsArrayLength = 0;
                    Count = 0;
                }
            }
        }

        internal void HardReset(int capacity)
        {
            ClearFast();
            Capacity = capacity;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void EnsureCapacity(int min)
        {
            if (_itemsArrayLength >= min) return;
            int newCapacity = _itemsArrayLength == 0 ? 4 : _itemsArrayLength * 2;
            if ((uint)newCapacity > 2146435071U) newCapacity = 2146435071;
            if (newCapacity < min) newCapacity = min;
            Capacity = newCapacity;
        }
    }
}
