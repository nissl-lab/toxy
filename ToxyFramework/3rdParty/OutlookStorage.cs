/*
 * OutlookStorage - Reads outlook msg file without Outlook object model - http://www.iwantedue.com
 * Copyright (C) 2008 David Ewen
 *
 * == BEGIN LICENSE ==
 *
 * Licensed under the terms of following license:
 *
 *  - The Code Project Open License 1.02 or later (the "CPOL")
 *    http://www.codeproject.com/info/cpol10.aspx
 *
 * == END LICENSE ==
 *
 * This file defines the OutlookStorage class used to read an outlook msg file.
 */

using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Reflection;

namespace iwantedue
{
    public class OutlookStorage: IDisposable
    {
        #region CLZF (This Region Has A Seperate Licence)

        /*
         * Copyright (c) 2005 Oren J. Maurice <oymaurice@hazorea.org.il>
         * 
         * Redistribution and use in source and binary forms, with or without modifica-
         * tion, are permitted provided that the following conditions are met:
         * 
         *   1.  Redistributions of source code must retain the above copyright notice,
         *       this list of conditions and the following disclaimer.
         * 
         *   2.  Redistributions in binary form must reproduce the above copyright
         *       notice, this list of conditions and the following disclaimer in the
         *       documentation and/or other materials provided with the distribution.
         * 
         *   3.  The name of the author may not be used to endorse or promote products
         *       derived from this software without specific prior written permission.
         * 
         * THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED
         * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MER-
         * CHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO
         * EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPE-
         * CIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
         * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
         * OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
         * WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTH-
         * ERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
         * OF THE POSSIBILITY OF SUCH DAMAGE.
         *
         * Alternatively, the contents of this file may be used under the terms of
         * the GNU General Public License version 2 (the "GPL"), in which case the
         * provisions of the GPL are applicable instead of the above. If you wish to
         * allow the use of your version of this file only under the terms of the
         * GPL and not to allow others to use your version of this file under the
         * BSD license, indicate your decision by deleting the provisions above and
         * replace them with the notice and other provisions required by the GPL. If
         * you do not delete the provisions above, a recipient may use your version
         * of this file under either the BSD or the GPL.
         */

        /// <summary>
        /// Summary description for CLZF.
        /// </summary>
        public class CLZF
        {
            /*
             This program is free software; you can redistribute it and/or modify
             it under the terms of the GNU General Public License as published by
             the Free Software Foundation; either version 2 of the License, or
             (at your option) any later version.

             You should have received a copy of the GNU General Public License
             along with this program; if not, write to the Free Software Foundation,
             Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA
            */

            /*
             * Prebuffered bytes used in RTF-compressed format (found them in RTFLIB32.LIB)
            */
            static byte[] COMPRESSED_RTF_PREBUF;
            static string prebuf = "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" +
                "{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " +
                "\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier" +
                "{\\colortbl\\red0\\green0\\blue0\n\r\\par " +
                "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";

            /* The lookup table used in the CRC32 calculation */
            static uint[] CRC32_TABLE =
	        {
		        0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA, 0x076DC419,
		        0x706AF48F, 0xE963A535, 0x9E6495A3, 0x0EDB8832, 0x79DCB8A4,
		        0xE0D5E91E, 0x97D2D988, 0x09B64C2B, 0x7EB17CBD, 0xE7B82D07,
		        0x90BF1D91, 0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE,
		        0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7, 0x136C9856,
		        0x646BA8C0, 0xFD62F97A, 0x8A65C9EC, 0x14015C4F, 0x63066CD9,
		        0xFA0F3D63, 0x8D080DF5, 0x3B6E20C8, 0x4C69105E, 0xD56041E4,
		        0xA2677172, 0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
		        0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940, 0x32D86CE3,
		        0x45DF5C75, 0xDCD60DCF, 0xABD13D59, 0x26D930AC, 0x51DE003A,
		        0xC8D75180, 0xBFD06116, 0x21B4F4B5, 0x56B3C423, 0xCFBA9599,
		        0xB8BDA50F, 0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924,
		        0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D, 0x76DC4190,
		        0x01DB7106, 0x98D220BC, 0xEFD5102A, 0x71B18589, 0x06B6B51F,
		        0x9FBFE4A5, 0xE8B8D433, 0x7807C9A2, 0x0F00F934, 0x9609A88E,
		        0xE10E9818, 0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
		        0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E, 0x6C0695ED,
		        0x1B01A57B, 0x8208F4C1, 0xF50FC457, 0x65B0D9C6, 0x12B7E950,
		        0x8BBEB8EA, 0xFCB9887C, 0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3,
		        0xFBD44C65, 0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2,
		        0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB, 0x4369E96A,
		        0x346ED9FC, 0xAD678846, 0xDA60B8D0, 0x44042D73, 0x33031DE5,
		        0xAA0A4C5F, 0xDD0D7CC9, 0x5005713C, 0x270241AA, 0xBE0B1010,
		        0xC90C2086, 0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
		        0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4, 0x59B33D17,
		        0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD, 0xEDB88320, 0x9ABFB3B6,
		        0x03B6E20C, 0x74B1D29A, 0xEAD54739, 0x9DD277AF, 0x04DB2615,
		        0x73DC1683, 0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8,
		        0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1, 0xF00F9344,
		        0x8708A3D2, 0x1E01F268, 0x6906C2FE, 0xF762575D, 0x806567CB,
		        0x196C3671, 0x6E6B06E7, 0xFED41B76, 0x89D32BE0, 0x10DA7A5A,
		        0x67DD4ACC, 0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
		        0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDFF252, 0xD1BB67F1,
		        0xA6BC5767, 0x3FB506DD, 0x48B2364B, 0xD80D2BDA, 0xAF0A1B4C,
		        0x36034AF6, 0x41047A60, 0xDF60EFC3, 0xA867DF55, 0x316E8EEF,
		        0x4669BE79, 0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236,
		        0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F, 0xC5BA3BBE,
		        0xB2BD0B28, 0x2BB45A92, 0x5CB36A04, 0xC2D7FFA7, 0xB5D0CF31,
		        0x2CD99E8B, 0x5BDEAE1D, 0x9B64C2B0, 0xEC63F226, 0x756AA39C,
		        0x026D930A, 0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
		        0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38, 0x92D28E9B,
		        0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21, 0x86D3D2D4, 0xF1D4E242,
		        0x68DDB3F8, 0x1FDA836E, 0x81BE16CD, 0xF6B9265B, 0x6FB077E1,
		        0x18B74777, 0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C,
		        0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45, 0xA00AE278,
		        0xD70DD2EE, 0x4E048354, 0x3903B3C2, 0xA7672661, 0xD06016F7,
		        0x4969474D, 0x3E6E77DB, 0xAED16A4A, 0xD9D65ADC, 0x40DF0B66,
		        0x37D83BF0, 0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
		        0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6, 0xBAD03605,
		        0xCDD70693, 0x54DE5729, 0x23D967BF, 0xB3667A2E, 0xC4614AB8,
		        0x5D681B02, 0x2A6F2B94, 0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B,
		        0x2D02EF8D
	        };

            /*
             * Calculates the CRC32 of the given bytes.
             * The CRC32 calculation is similar to the standard one as demonstrated
             * in RFC 1952, but with the inversion (before and after the calculation)
             * ommited.
             * 
             * @param buf the byte array to calculate CRC32 on
             * @param off the offset within buf at which the CRC32 calculation will start
             * @param len the number of bytes on which to calculate the CRC32
             * @return the CRC32 value.
             */
            static public int calculateCRC32(byte[] buf, int off, int len)
            {
                uint c = 0;
                int end = off + len;
                for (int i = off; i < end; i++)
                {
                    //!!!!        c = CRC32_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
                    c = CRC32_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >> 8);
                }
                return (int)c;
            }

            /*
                 * Returns an unsigned 32-bit value from little-endian ordered bytes.
                 *
                 * @param   buf a byte array from which byte values are taken
                 * @param   offset the offset within buf from which byte values are taken
                 * @return  an unsigned 32-bit value as a long.
            */
            public static long getU32(byte[] buf, int offset)
            {
                return ((buf[offset] & 0xFF) | ((buf[offset + 1] & 0xFF) << 8) | ((buf[offset + 2] & 0xFF) << 16) | ((buf[offset + 3] & 0xFF) << 24)) & 0x00000000FFFFFFFFL;
            }

            /*
             * Returns an unsigned 8-bit value from a byte array.
             *
             * @param   buf a byte array from which byte value is taken
             * @param   offset the offset within buf from which byte value is taken
             * @return  an unsigned 8-bit value as an int.
             */
            public static int getU8(byte[] buf, int offset)
            {
                return buf[offset] & 0xFF;
            }

            /*
              * Decompresses compressed-RTF data.
              *
              * @param   src the compressed-RTF data bytes
              * @return  an array containing the decompressed bytes.
              * @throws  IllegalArgumentException if src does not contain valid                                                                                                                                            *          compressed-RTF bytes.
            */
            public static byte[] decompressRTF(byte[] src)
            {
                byte[] dst; // destination for uncompressed bytes
                int inPos = 0; // current position in src array
                int outPos = 0; // current position in dst array

                COMPRESSED_RTF_PREBUF = System.Text.Encoding.ASCII.GetBytes(prebuf);

                // get header fields (as defined in RTFLIB.H)
                if (src == null || src.Length < 16)
                    throw new Exception("Invalid compressed-RTF header");

                int compressedSize = (int)getU32(src, inPos);
                inPos += 4;
                int uncompressedSize = (int)getU32(src, inPos);
                inPos += 4;
                int magic = (int)getU32(src, inPos);
                inPos += 4;
                int crc32 = (int)getU32(src, inPos);
                inPos += 4;

                if (compressedSize != src.Length - 4) // check size excluding the size field itself
                    throw new Exception("compressed-RTF data size mismatch");

                if (crc32 != calculateCRC32(src, 16, src.Length - 16))
                    throw new Exception("compressed-RTF CRC32 failed");

                // process the data
                if (magic == 0x414c454d)
                { // magic number that identifies the stream as a uncompressed stream
                    dst = new byte[uncompressedSize];
                    Array.Copy(src, inPos, dst, outPos, uncompressedSize); // just copy it as it is
                }
                else if (magic == 0x75465a4c)
                { // magic number that identifies the stream as a compressed stream
                    dst = new byte[COMPRESSED_RTF_PREBUF.Length + uncompressedSize];
                    Array.Copy(COMPRESSED_RTF_PREBUF, 0, dst, 0, COMPRESSED_RTF_PREBUF.Length);
                    outPos = COMPRESSED_RTF_PREBUF.Length;
                    int flagCount = 0;
                    int flags = 0;
                    while (outPos < dst.Length)
                    {
                        // each flag byte flags 8 literals/references, 1 per bit
                        flags = (flagCount++ % 8 == 0) ? getU8(src, inPos++) : flags >> 1;
                        if ((flags & 1) == 1)
                        { // each flag bit is 1 for reference, 0 for literal
                            int offset = getU8(src, inPos++);
                            int length = getU8(src, inPos++);
                            //!!!!!!!!!            offset = (offset << 4) | (length >>> 4); // the offset relative to block start
                            offset = (offset << 4) | (length >> 4); // the offset relative to block start
                            length = (length & 0xF) + 2; // the number of bytes to copy
                            // the decompression buffer is supposed to wrap around back
                            // to the beginning when the end is reached. we save the
                            // need for such a buffer by pointing straight into the data
                            // buffer, and simulating this behaviour by modifying the
                            // pointers appropriately.
                            offset = (outPos / 4096) * 4096 + offset;
                            if (offset >= outPos) // take from previous block
                                offset -= 4096;
                            // note: can't use System.arraycopy, because the referenced
                            // bytes can cross through the current out position.
                            int end = offset + length;
                            while (offset < end)
                                dst[outPos++] = dst[offset++];
                        }
                        else
                        { // literal
                            dst[outPos++] = src[inPos++];
                        }
                    }
                    // copy it back without the prebuffered data
                    src = dst;
                    dst = new byte[uncompressedSize];
                    Array.Copy(src, COMPRESSED_RTF_PREBUF.Length, dst, 0, uncompressedSize);
                }
                else
                { // unknown magic number
                    throw new Exception("Unknown compression type (magic number " + magic + ")");
                }

                return dst;
            }
        }

        #endregion

        #region NativeMethods

        protected class NativeMethods
        {
            [DllImport("kernel32.dll")]
            static extern IntPtr GlobalLock(IntPtr hMem);

            [DllImport("ole32.DLL")]
            public static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out ILockBytes ppLkbyt);

            //[DllImport("ole32.DLL", CharSet = CharSet.Auto, PreserveSig = false)]
            //public static extern IntPtr GetHGlobalFromILockBytes(ILockBytes pLockBytes);

            [DllImport("ole32.DLL")]
            public static extern int StgIsStorageILockBytes(ILockBytes plkbyt);

            [DllImport("ole32.DLL")]
            public static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, STGM grfMode, uint reserved, out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            public static extern void StgOpenStorageOnILockBytes(ILockBytes plkbyt, IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            public static extern int StgIsStorageFile([MarshalAs(UnmanagedType.LPWStr)] string wcsName);

            [DllImport("ole32.DLL")]
            public static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string wcsName, IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, int reserved, out IStorage ppstgOpen);

            [ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface ILockBytes
            {
                void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbRead);
                void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);
                void Flush();
                void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);
                void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
                void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
                void Stat([Out]out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
            }

            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
            public interface IStorage
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                ComTypes.IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                ComTypes.IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.U4)] int reserved);
                void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIIDExclude, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);
                void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName, [In, MarshalAs(UnmanagedType.U4)] int grfFlags);
                void Commit(int grfCommitFlags);
                void Revert();
                void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2, [In, MarshalAs(UnmanagedType.U4)] int reserved3, [MarshalAs(UnmanagedType.Interface)] out IEnumSTATSTG ppVal);
                void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);
                void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);
                void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] System.Runtime.InteropServices.ComTypes.FILETIME pctime, [In] System.Runtime.InteropServices.ComTypes.FILETIME patime, [In] System.Runtime.InteropServices.ComTypes.FILETIME pmtime);
                void SetClass([In] ref Guid clsid);
                void SetStateBits(int grfStateBits, int grfMask);
                void Stat([Out]out System.Runtime.InteropServices.ComTypes.STATSTG pStatStg, int grfStatFlag);
            }

            [ComImport, Guid("0000000D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IEnumSTATSTG
            {
                void Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt, out uint pceltFetched);
                void Skip(uint celt);
                void Reset();
                [return: MarshalAs(UnmanagedType.Interface)]
                IEnumSTATSTG Clone();
            }

            public enum STGM : int
            {
                DIRECT = 0x00000000,
                TRANSACTED = 0x00010000,
                SIMPLE = 0x08000000,
                READ = 0x00000000,
                WRITE = 0x00000001,
                READWRITE = 0x00000002,
                SHARE_DENY_NONE = 0x00000040,
                SHARE_DENY_READ = 0x00000030,
                SHARE_DENY_WRITE = 0x00000020,
                SHARE_EXCLUSIVE = 0x00000010,
                PRIORITY = 0x00040000,
                DELETEONRELEASE = 0x04000000,
                NOSCRATCH = 0x00100000,
                CREATE = 0x00001000,
                CONVERT = 0x00020000,
                FAILIFTHERE = 0x00000000,
                NOSNAPSHOT = 0x00200000,
                DIRECT_SWMR = 0x00400000
            }

            public const ushort PT_UNSPECIFIED = 0; /* (Reserved for interface use) type doesn't matter to caller */
            public const ushort PT_NULL = 1;        /* NULL property value */
            public const ushort PT_I2 = 2;          /* Signed 16-bit value */
            public const ushort PT_LONG = 3;        /* Signed 32-bit value */
            public const ushort PT_R4 = 4;          /* 4-byte floating point */
            public const ushort PT_DOUBLE = 5;      /* Floating point double */
            public const ushort PT_CURRENCY = 6;    /* Signed 64-bit int (decimal w/    4 digits right of decimal pt) */
            public const ushort PT_APPTIME = 7;     /* Application time */
            public const ushort PT_ERROR = 10;      /* 32-bit error value */
            public const ushort PT_BOOLEAN = 11;    /* 16-bit boolean (non-zero true) */
            public const ushort PT_OBJECT = 13;     /* Embedded object in a property */
            public const ushort PT_I8 = 20;         /* 8-byte signed integer */
            public const ushort PT_STRING8 = 30;    /* Null terminated 8-bit character string */
            public const ushort PT_UNICODE = 31;    /* Null terminated Unicode string */
            public const ushort PT_SYSTIME = 64;    /* FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601 */
            public const ushort PT_CLSID = 72;      /* OLE GUID */
            public const ushort PT_BINARY = 258;    /* Uninterpreted (counted byte array) */

            public static IStorage CloneStorage(IStorage source, bool closeSource)
            {
                NativeMethods.IStorage memoryStorage = null;
                NativeMethods.ILockBytes memoryStorageBytes = null;
                try
                {
                    //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    NativeMethods.StgCreateDocfileOnILockBytes(memoryStorageBytes, NativeMethods.STGM.CREATE | NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, out memoryStorage);

                    //copy the source storage into the new storage
                    source.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                    memoryStorageBytes.Flush();
                    memoryStorage.Commit(0);

                    //ensure memory is released
                    ReferenceManager.AddItem(memoryStorage);
                }
                catch
                {
                    if (memoryStorage != null)
                    {
                        Marshal.ReleaseComObject(memoryStorage);
                    }
                }
                finally
                {
                    if (memoryStorageBytes != null)
                    {
                        Marshal.ReleaseComObject(memoryStorageBytes);
                    }

                    if (closeSource)
                    {
                        Marshal.ReleaseComObject(source);
                    }
                }

                return memoryStorage;
            }
        }

        #endregion

        #region ReferenceManager

        private class ReferenceManager
        {
            public static void AddItem(object track)
            {
                lock (instance)
                {
                    if (!instance.trackingObjects.Contains(track))
                    {
                        instance.trackingObjects.Add(track);
                    }
                }
            }

            public static void RemoveItem(object track)
            {
                lock (instance)
                {
                    if (instance.trackingObjects.Contains(track))
                    {
                        instance.trackingObjects.Remove(track);
                    }
                }
            }

            private static ReferenceManager instance = new ReferenceManager();

            private List<object> trackingObjects = new List<object>();

            private ReferenceManager() { }

            ~ReferenceManager()
            {
                foreach (object trackingObject in trackingObjects)
                {
                    Marshal.ReleaseComObject(trackingObject);
                }
            }
        }

        #endregion

        #region Nested Classes

        public enum RecipientType
        {
            To,
            CC,
            Unknown
        }

        public class Recipient : OutlookStorage
        {
            #region Property(s)

            /// <summary>
            /// Gets the display name.
            /// </summary>
            /// <value>The display name.</value>
            public string DisplayName
            {
                get { return this.GetMapiPropertyString(OutlookStorage.PR_DISPLAY_NAME); }
            }

            /// <summary>
            /// Gets the recipient email.
            /// </summary>
            /// <value>The recipient email.</value>
            public string Email
            {
                get
                {
                    string email = this.GetMapiPropertyString(OutlookStorage.PR_EMAIL);
                    if (String.IsNullOrEmpty(email))
                    {
                        email = this.GetMapiPropertyString(OutlookStorage.PR_EMAIL_2);
                    }
                    return email;
                }
            }

            /// <summary>
            /// Gets the recipient type.
            /// </summary>
            /// <value>The recipient type.</value>
            public RecipientType Type
            {
                get
                {
                    int recipientType = this.GetMapiPropertyInt32(OutlookStorage.PR_RECIPIENT_TYPE);
                    switch (recipientType)
                    {
                        case OutlookStorage.MAPI_TO:
                            return RecipientType.To;

                        case OutlookStorage.MAPI_CC:
                            return RecipientType.CC;
                    }
                    return RecipientType.Unknown;
                }
            }

            #endregion

            #region Constructor(s)

            /// <summary>
            /// Initializes a new instance of the <see cref="Recipient"/> class.
            /// </summary>
            /// <param name="message">The message.</param>
            public Recipient(OutlookStorage message)
                : base(message.storage)
            {
                GC.SuppressFinalize(message);
                this.propHeaderSize = OutlookStorage.PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP;
            }

            #endregion
        }

        public class Attachment : OutlookStorage
        {
            #region Property(s)

            /// <summary>
            /// Gets the filename.
            /// </summary>
            /// <value>The filename.</value>
            public string Filename
            {
                get
                {
                    string filename = this.GetMapiPropertyString(OutlookStorage.PR_ATTACH_LONG_FILENAME);
                    if (String.IsNullOrEmpty(filename))
                    {
                        filename = this.GetMapiPropertyString(OutlookStorage.PR_ATTACH_FILENAME);
                    }
                    if (String.IsNullOrEmpty(filename))
                    {
                        filename = this.GetMapiPropertyString(OutlookStorage.PR_DISPLAY_NAME);
                    }
                    return filename;
                }
            }

            /// <summary>
            /// Gets the data.
            /// </summary>
            /// <value>The data.</value>
            public byte[] Data
            {
                get { return this.GetMapiPropertyBytes(OutlookStorage.PR_ATTACH_DATA); }
            }

            /// <summary>
            /// Gets the content id.
            /// </summary>
            /// <value>The content id.</value>
            public string ContentId
            {
                get { return this.GetMapiPropertyString(OutlookStorage.PR_ATTACH_CONTENT_ID); }
            }

            /// <summary>
            /// Gets the rendering posisiton.
            /// </summary>
            /// <value>The rendering posisiton.</value>
            public int RenderingPosisiton
            {
                get { return this.GetMapiPropertyInt32(OutlookStorage.PR_RENDERING_POSITION); }
            }

            #endregion

            #region Constructor(s)

            /// <summary>
            /// Initializes a new instance of the <see cref="Attachment"/> class.
            /// </summary>
            /// <param name="message">The message.</param>
            public Attachment(OutlookStorage message)
                : base(message.storage)
            {
                GC.SuppressFinalize(message);
                this.propHeaderSize = OutlookStorage.PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP;
            }

            #endregion
        }

        public class Message : OutlookStorage
        {
            #region Property(s)

            /// <summary>
            /// Gets the list of recipients in the outlook message.
            /// </summary>
            /// <value>The list of recipients in the outlook message.</value>
            public List<Recipient> Recipients
            {
                get { return this.recipients; }
            }
            private List<Recipient> recipients = new List<Recipient>();

            /// <summary>
            /// Gets the list of attachments in the outlook message.
            /// </summary>
            /// <value>The list of attachments in the outlook message.</value>
            public List<Attachment> Attachments
            {
                get { return this.attachments; }
            }
            private List<Attachment> attachments = new List<Attachment>();

            /// <summary>
            /// Gets the list of sub messages in the outlook message.
            /// </summary>
            /// <value>The list of sub messages in the outlook message.</value>
            public List<Message> Messages
            {
                get { return this.messages; }
            }
            private List<Message> messages = new List<Message>();

            /// <summary>
            /// Gets the display value of the contact that sent the email.
            /// </summary>
            /// <value>The display value of the contact that sent the email.</value>
            public String From
            {
                get { return this.GetMapiPropertyString(OutlookStorage.PR_SENDER_NAME); }
            }

            /// <summary>
            /// Gets the subject of the outlook message.
            /// </summary>
            /// <value>The subject of the outlook message.</value>
            public String Subject
            {
                get { return this.GetMapiPropertyString(OutlookStorage.PR_SUBJECT); }
            }

            /// <summary>
            /// Gets the body of the outlook message in plain text format.
            /// </summary>
            /// <value>The body of the outlook message in plain text format.</value>
            public String BodyText
            {
                get { return this.GetMapiPropertyString(OutlookStorage.PR_BODY); }
            }

            /// <summary>
            /// Gets the body of the outlook message in RTF format.
            /// </summary>
            /// <value>The body of the outlook message in RTF format.</value>
            public String BodyRTF
            {
                get
                {
                    //get value for the RTF compressed MAPI property
                    byte[] rtfBytes = this.GetMapiPropertyBytes(OutlookStorage.PR_RTF_COMPRESSED);

                    //return null if no property value exists
                    if (rtfBytes == null || rtfBytes.Length == 0)
                    {
                        return null;
                    }

                    //decompress the rtf value
                    rtfBytes = CLZF.decompressRTF(rtfBytes);

                    //encode the rtf value as an ascii string and return
                    return Encoding.ASCII.GetString(rtfBytes);
                }
            }

            #endregion

            #region Constructor(s)

            /// <summary>
            /// Initializes a new instance of the <see cref="Message"/> class from a msg file.
            /// </summary>
            /// <param name="filename">The msg file to load.</param>
            public Message(string msgfile): base(msgfile) { }

            /// <summary>
            /// Initializes a new instance of the <see cref="Message"/> class from a <see cref="Stream"/> containing an IStorage.
            /// </summary>
            /// <param name="storageStream">The <see cref="Stream"/> containing an IStorage.</param>
            public Message(Stream storageStream): base(storageStream) { }

            /// <summary>
            /// Initializes a new instance of the <see cref="Message"/> class on the specified <see cref="NativeMethods.IStorage"/>.
            /// </summary>
            /// <param name="storage">The storage to create the <see cref="Message"/> on.</param>
            private Message(NativeMethods.IStorage storage)
                : base(storage)
            {
                this.propHeaderSize = OutlookStorage.PROPERTIES_STREAM_HEADER_TOP;
            }

            #endregion

            #region Methods(LoadStorage)

            /// <summary>
            /// Processes sub storages on the specified storage to capture attachment and recipient data.
            /// </summary>
            /// <param name="storage">The storage to check for attachment and recipient data.</param>
            protected override void LoadStorage(NativeMethods.IStorage storage)
            {
                base.LoadStorage(storage);

                foreach (ComTypes.STATSTG storageStat in this.subStorageStatistics.Values)
                {
                    //element is a storage. get it and add its statistics object to the sub storage dictionary
                    NativeMethods.IStorage subStorage = this.storage.OpenStorage(storageStat.pwcsName, IntPtr.Zero, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0);

                    //run specific load method depending on sub storage name prefix
                    if (storageStat.pwcsName.StartsWith(OutlookStorage.RECIP_STORAGE_PREFIX))
                    {
                        Recipient recipient = new Recipient(new OutlookStorage(subStorage));
                        this.recipients.Add(recipient);
                    }
                    else if (storageStat.pwcsName.StartsWith(OutlookStorage.ATTACH_STORAGE_PREFIX))
                    {
                        this.LoadAttachmentStorage(subStorage);
                    }
                    else
                    {
                        //release sub storage
                        Marshal.ReleaseComObject(subStorage);
                    }
                }
            }

            /// <summary>
            /// Loads the attachment data out of the specified storage.
            /// </summary>
            /// <param name="storage">The attachment storage.</param>
            private void LoadAttachmentStorage(NativeMethods.IStorage storage)
            {
                //create attachment from attachment storage
                Attachment attachment = new Attachment(new OutlookStorage(storage));

                //if attachment is a embeded msg handle differently than an normal attachment
                int attachMethod = attachment.GetMapiPropertyInt32(OutlookStorage.PR_ATTACH_METHOD);
                if (attachMethod == OutlookStorage.ATTACH_EMBEDDED_MSG)
                {
                    //create new Message and set parent and header size
                    Message subMsg = new Message(attachment.GetMapiProperty(OutlookStorage.PR_ATTACH_DATA) as NativeMethods.IStorage);
                    subMsg.parentMessage = this;
                    subMsg.propHeaderSize = OutlookStorage.PROPERTIES_STREAM_HEADER_EMBEDED;

                    //add to messages list
                    this.messages.Add(subMsg);
                }
                else
                {
                    //add attachment to attachment list
                    this.attachments.Add(attachment);
                }
            }

            #endregion

            #region Methods(Save)

            /// <summary>
            /// Saves this <see cref="Message"/> to the specified file name.
            /// </summary>
            /// <param name="fileName">Name of the file.</param>
            public void Save(string fileName)
            {
                FileStream saveFileStream = File.Open(fileName, FileMode.Create, FileAccess.ReadWrite);
                this.Save(saveFileStream);
                saveFileStream.Close();
            }

            /// <summary>
            /// Saves this <see cref="Message"/> to the specified stream.
            /// </summary>
            /// <param name="stream">The stream to save to.</param>
            public void Save(Stream stream)
            {
                //get statistics for stream 
                OutlookStorage saveMsg = this;

                byte[] memoryStorageContent;
                NativeMethods.IStorage memoryStorage = null;
                NativeMethods.IStorage nameIdStorage = null;
                NativeMethods.IStorage nameIdSourceStorage = null;
                NativeMethods.ILockBytes memoryStorageBytes = null;
                try
                {
                    //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    NativeMethods.StgCreateDocfileOnILockBytes(memoryStorageBytes, NativeMethods.STGM.CREATE | NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, out memoryStorage);

                    //copy the save storage into the new storage
                    saveMsg.storage.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                    memoryStorageBytes.Flush();
                    memoryStorage.Commit(0);

                    //if not the top parent then the name id mapping needs to be copied from top parent to this message and the property stream header needs to be padded by 8 bytes
                    if (!this.IsTopParent)
                    {
                        //create a new name id storage and get the source name id storage to copy from
                        nameIdStorage = memoryStorage.CreateStorage(OutlookStorage.NAMEID_STORAGE, NativeMethods.STGM.CREATE | NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, 0);
                        nameIdSourceStorage = this.TopParent.storage.OpenStorage(OutlookStorage.NAMEID_STORAGE, IntPtr.Zero, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0);

                        //copy the name id storage from the parent to the new name id storage
                        nameIdSourceStorage.CopyTo(0, null, IntPtr.Zero, nameIdStorage);

                        //get the property bytes for the storage being copied
                        byte[] props = saveMsg.GetStreamBytes(OutlookStorage.PROPERTIES_STREAM);

                        //create new array to store a copy of the properties that is 8 bytes larger than the old so the header can be padded
                        byte[] newProps = new byte[props.Length + 8];

                        //insert 8 null bytes from index 24 to 32. this is because a top level object property header requires a 32 byte header
                        Buffer.BlockCopy(props, 0, newProps, 0, 24);
                        Buffer.BlockCopy(props, 24, newProps, 32, props.Length - 24);

                        //remove the copied prop bytes so it can be replaced with the padded version
                        memoryStorage.DestroyElement(OutlookStorage.PROPERTIES_STREAM);

                        //create the property stream again and write in the padded version
                        ComTypes.IStream propStream = memoryStorage.CreateStream(OutlookStorage.PROPERTIES_STREAM, NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, 0);
                        propStream.Write(newProps, newProps.Length, IntPtr.Zero);
                    }

                    //commit changes to the storage
                    memoryStorage.Commit(0);
                    memoryStorageBytes.Flush();

                    //get the STATSTG of the ILockBytes to determine how many bytes were written to it
                    ComTypes.STATSTG memoryStorageBytesStat;
                    memoryStorageBytes.Stat(out memoryStorageBytesStat, 1);

                    //read the bytes into a managed byte array
                    memoryStorageContent = new byte[memoryStorageBytesStat.cbSize];
                    memoryStorageBytes.ReadAt(0, memoryStorageContent, memoryStorageContent.Length, null);

                    //write storage bytes to stream
                    stream.Write(memoryStorageContent, 0, memoryStorageContent.Length);
                }
                finally
                {
                    if (nameIdSourceStorage != null)
                    {
                        Marshal.ReleaseComObject(nameIdSourceStorage);
                    }

                    if (memoryStorage != null)
                    {
                        Marshal.ReleaseComObject(memoryStorage);
                    }

                    if (memoryStorageBytes != null)
                    {
                        Marshal.ReleaseComObject(memoryStorageBytes);
                    }
                }
            }

            #endregion

            #region Methods(Disposing)

            protected override void Disposing()
            {
                //dispose sub storages
                foreach (OutlookStorage subMsg in this.messages)
                {
                    subMsg.Dispose();
                }

                //dispose sub storages
                foreach (OutlookStorage recip in this.recipients)
                {
                    recip.Dispose();
                }

                //dispose sub storages
                foreach (OutlookStorage attach in this.attachments)
                {
                    attach.Dispose();
                }
            }

            #endregion
        }

        #endregion

        #region Constants

        //attachment constants
        private const string ATTACH_STORAGE_PREFIX = "__attach_version1.0_#";
        private const string PR_ATTACH_FILENAME = "3704";
        private const string PR_ATTACH_LONG_FILENAME = "3707";
        private const string PR_ATTACH_DATA = "3701";
        private const string PR_ATTACH_METHOD = "3705";
        private const string PR_RENDERING_POSITION = "370B";
        private const string PR_ATTACH_CONTENT_ID = "3712";
        private const int ATTACH_BY_VALUE = 1;
        private const int ATTACH_EMBEDDED_MSG = 5;

        //recipient constants
        private const string RECIP_STORAGE_PREFIX = "__recip_version1.0_#";
        private const string PR_DISPLAY_NAME = "3001";
        private const string PR_EMAIL = "39FE";
        private const string PR_EMAIL_2 = "403E"; //not sure why but email address is in this property sometimes cant find any documentation on it
        private const string PR_RECIPIENT_TYPE = "0C15";
        private const int MAPI_TO = 1;
        private const int MAPI_CC = 2;

        //msg constants
        private const string PR_SUBJECT = "0037";
        private const string PR_BODY = "1000";
        private const string PR_RTF_COMPRESSED = "1009";
        private const string PR_SENDER_NAME = "0C1A";

        //property stream constants
        private const string PROPERTIES_STREAM = "__properties_version1.0";
        private const int PROPERTIES_STREAM_HEADER_TOP = 32;
        private const int PROPERTIES_STREAM_HEADER_EMBEDED = 24;
        private const int PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP = 8;

        //name id storage name in root storage
        private const string NAMEID_STORAGE = "__nameid_version1.0";

        #endregion

        #region Property(s)

        /// <summary>
        /// Gets the top level outlook message from a sub message at any level.
        /// </summary>
        /// <value>The top level outlook message.</value>
        private OutlookStorage TopParent
        {
            get
            {
                if (this.parentMessage != null)
                {
                    return this.parentMessage.TopParent;
                }
                return this;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is the top level outlook message.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is the top level outlook message; otherwise, <c>false</c>.
        /// </value>
        private bool IsTopParent
        {
            get
            {
                if (this.parentMessage != null)
                {
                    return false;
                }
                return true;
            }
        }

        /// <summary>
        /// The IStorage associated with this instance.
        /// </summary>
        private NativeMethods.IStorage storage;

        /// <summary>
        /// Header size of the property stream in the IStorage associated with this instance.
        /// </summary>
        private int propHeaderSize = OutlookStorage.PROPERTIES_STREAM_HEADER_TOP;

        /// <summary>
        /// A reference to the parent message that this message may belong to.
        /// </summary>
        private OutlookStorage parentMessage = null;

        /// <summary>
        /// The statistics for all streams in the IStorage associated with this instance.
        /// </summary>
        public Dictionary<string, ComTypes.STATSTG> streamStatistics = new Dictionary<string, ComTypes.STATSTG>();

        /// <summary>
        /// The statistics for all storgages in the IStorage associated with this instance.
        /// </summary>
        public Dictionary<string, ComTypes.STATSTG> subStorageStatistics = new Dictionary<string, ComTypes.STATSTG>();

        /// <summary>
        /// Indicates wether this instance has been disposed.
        /// </summary>
        private bool disposed = false;

        #endregion

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookStorage"/> class from a file.
        /// </summary>
        /// <param name="storageFilePath">The file to load.</param>
        private OutlookStorage(string storageFilePath)
        {
            //ensure provided file is an IStorage
            if (NativeMethods.StgIsStorageFile(storageFilePath) != 0)
            {
                throw new ArgumentException("The provided file is not a valid IStorage", "storageFilePath");
            }

            //open and load IStorage from file
            NativeMethods.IStorage fileStorage;
            NativeMethods.StgOpenStorage(storageFilePath, null, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out fileStorage);
            this.LoadStorage(fileStorage);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookStorage"/> class from a <see cref="Stream"/> containing an IStorage.
        /// </summary>
        /// <param name="storageStream">The <see cref="Stream"/> containing an IStorage.</param>
        private OutlookStorage(Stream storageStream)
        {
            NativeMethods.IStorage memoryStorage = null;
            NativeMethods.ILockBytes memoryStorageBytes = null;
            try
            {
                //read stream into buffer
                byte[] buffer = new byte[storageStream.Length];
                storageStream.Read(buffer, 0, buffer.Length);

                //create a ILockBytes (unmanaged byte array) and write buffer into it
                NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                memoryStorageBytes.WriteAt(0, buffer, buffer.Length, null);

                //ensure provided stream data is an IStorage
                if (NativeMethods.StgIsStorageILockBytes(memoryStorageBytes) != 0)
                {
                    throw new ArgumentException("The provided stream is not a valid IStorage", "storageStream");
                }

                //open and load IStorage on the ILockBytes
                NativeMethods.StgOpenStorageOnILockBytes(memoryStorageBytes, null, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out memoryStorage);
                this.LoadStorage(memoryStorage);
            }
            catch
            {
                if (memoryStorage != null)
                {
                    Marshal.ReleaseComObject(memoryStorage);
                }
            }
            finally
            {
                if (memoryStorageBytes != null)
                {
                    Marshal.ReleaseComObject(memoryStorageBytes);
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookStorage"/> class on the specified <see cref="NativeMethods.IStorage"/>.
        /// </summary>
        /// <param name="storage">The storage to create the <see cref="OutlookStorage"/> on.</param>
        private OutlookStorage(NativeMethods.IStorage storage)
        {
            this.LoadStorage(storage);
        }

        /// <summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="OutlookStorage"/> is reclaimed by garbage collection.
        /// </summary>
        ~OutlookStorage()
        {
            this.Dispose();
        }

        #endregion

        #region Methods(LoadStorage)

        /// <summary>
        /// Processes sub streams and storages on the specified storage.
        /// </summary>
        /// <param name="storage">The storage to get sub streams and storages for.</param>
        protected virtual void LoadStorage(NativeMethods.IStorage storage)
        {
            this.storage = storage;

            //ensures memory is released
            ReferenceManager.AddItem(this.storage);

            NativeMethods.IEnumSTATSTG storageElementEnum = null;
            try
            {
                //enum all elements of the storage
                storage.EnumElements(0, IntPtr.Zero, 0, out storageElementEnum);

                //iterate elements
                while (true)
                {
                    //get 1 element out of the com enumerator
                    uint elementStatCount;
                    ComTypes.STATSTG[] elementStats = new ComTypes.STATSTG[1];
                    storageElementEnum.Next(1, elementStats, out elementStatCount);

                    //break loop if element not retrieved
                    if (elementStatCount != 1)
                    {
                        break;
                    }

                    ComTypes.STATSTG elementStat = elementStats[0];
                    switch (elementStat.type)
                    {
                        case 1:
                            //element is a storage. add its statistics object to the storage dictionary
                            subStorageStatistics.Add(elementStat.pwcsName, elementStat);
                            break;

                        case 2:
                            //element is a stream. add its statistics object to the stream dictionary
                            streamStatistics.Add(elementStat.pwcsName, elementStat);
                            break;
                    }
                }
            }
            finally
            {
                //free memory
                if (storageElementEnum != null)
                {
                    Marshal.ReleaseComObject(storageElementEnum);
                }
            }
        }

        #endregion

        #region Methods(GetStreamBytes, GetStreamAsString)

        /// <summary>
        /// Gets the data in the specified stream as a byte array.
        /// </summary>
        /// <param name="streamName">Name of the stream to get data for.</param>
        /// <returns>A byte array containg the stream data.</returns>
        public byte[] GetStreamBytes(string streamName)
        {
            //get statistics for stream 
            ComTypes.STATSTG streamStatStg = this.streamStatistics[streamName];

            byte[] iStreamContent;
            ComTypes.IStream stream = null;
            try
            {
                //open stream from the storage
                stream = this.storage.OpenStream(streamStatStg.pwcsName, IntPtr.Zero, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE, 0);

                //read the stream into a managed byte array
                iStreamContent = new byte[streamStatStg.cbSize];
                stream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero);
            }
            finally
            {
                if (stream != null)
                {
                    Marshal.ReleaseComObject(stream);
                }
            }

            //return the stream bytes
            return iStreamContent;
        }

        /// <summary>
        /// Gets the data in the specified stream as a string using the specifed encoding to decode the stream data.
        /// </summary>
        /// <param name="streamName">Name of the stream to get string data for.</param>
        /// <param name="streamEncoding">The encoding to decode the stream data with.</param>
        /// <returns>The data in the specified stream as a string.</returns>
        public string GetStreamAsString(string streamName, Encoding streamEncoding)
        {
            StreamReader streamReader = new StreamReader(new MemoryStream(this.GetStreamBytes(streamName)), streamEncoding);
            string streamContent = streamReader.ReadToEnd();
            streamReader.Close();

            return streamContent;
        }

        #endregion

        #region Methods(GetMapiProperty)

        /// <summary>
        /// Gets the raw value of the MAPI property.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The raw value of the MAPI property.</returns>
        public object GetMapiProperty(string propIdentifier)
        {
            //try get prop value from stream or storage
            object propValue = this.GetMapiPropertyFromStreamOrStorage(propIdentifier);

            //if not found in stream or storage try get prop value from property stream
            if (propValue == null)
            {
                propValue = this.GetMapiPropertyFromPropertyStream(propIdentifier);
            }

            return propValue;
        }

        /// <summary>
        /// Gets the MAPI property value from a stream or storage in this storage.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The value of the MAPI property or null if not found.</returns>
        private object GetMapiPropertyFromStreamOrStorage(string propIdentifier)
        {
            //get list of stream and storage identifiers which map to properties
            List<string> propKeys = new List<string>();
            propKeys.AddRange(this.streamStatistics.Keys);
            propKeys.AddRange(this.subStorageStatistics.Keys);

            //determine if the property identifier is in a stream or sub storage
            string propTag = null;
            ushort propType = NativeMethods.PT_UNSPECIFIED;
            foreach (string propKey in propKeys)
            {
                if (propKey.StartsWith("__substg1.0_" + propIdentifier))
                {
                    propTag = propKey.Substring(12, 8);
                    propType = ushort.Parse(propKey.Substring(16, 4), System.Globalization.NumberStyles.HexNumber);
                    break;
                }
            }

            //depending on prop type use method to get property value
            string containerName = "__substg1.0_" + propTag;
            switch (propType)
            {
                case NativeMethods.PT_UNSPECIFIED:
                    return null;

                case NativeMethods.PT_STRING8:
                    return this.GetStreamAsString(containerName, Encoding.UTF8);

                case NativeMethods.PT_UNICODE:
                    return this.GetStreamAsString(containerName, Encoding.Unicode);

                case NativeMethods.PT_BINARY:
                    return this.GetStreamBytes(containerName);

                case NativeMethods.PT_OBJECT:
                    return NativeMethods.CloneStorage(this.storage.OpenStorage(containerName, IntPtr.Zero, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0), true);

                default:
                    throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
            }
        }

        /// <summary>
        /// Gets the MAPI property value from the property stream in this storage.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The value of the MAPI property or null if not found.</returns>
        private object GetMapiPropertyFromPropertyStream(string propIdentifier)
        {
            //if no property stream return null
            if (!this.streamStatistics.ContainsKey(OutlookStorage.PROPERTIES_STREAM))
            {
                return null;
            }

            //get the raw bytes for the property stream
            byte[] propBytes = this.GetStreamBytes(OutlookStorage.PROPERTIES_STREAM);

            //iterate over property stream in 16 byte chunks starting from end of header
            for (int i = this.propHeaderSize; i < propBytes.Length; i = i + 16)
            {
                //get property type located in the 1st and 2nd bytes as a unsigned short value
                ushort propType = BitConverter.ToUInt16(propBytes, i);

                //get property identifer located in 3nd and 4th bytes as a hexdecimal string
                byte[] propIdent = new byte[] { propBytes[i + 3], propBytes[i + 2] };
                string propIdentString = BitConverter.ToString(propIdent).Replace("-", "");

                //if this is not the property being gotten continue to next property
                if (propIdentString != propIdentifier)
                {
                    continue;
                }

                //depending on prop type use method to get property value
                switch (propType)
                {
                    case NativeMethods.PT_I2:
                        return BitConverter.ToInt16(propBytes, i + 8);

                    case NativeMethods.PT_LONG:
                        return BitConverter.ToInt32(propBytes, i + 8);

                    default:
                        throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
                }
            }

            //property not found return null
            return null;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a string.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The value of the MAPI property as a string.</returns>
        public string GetMapiPropertyString(string propIdentifier)
        {
            return this.GetMapiProperty(propIdentifier) as string;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a short.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The value of the MAPI property as a short.</returns>
        public Int16 GetMapiPropertyInt16(string propIdentifier)
        {
            return (Int16)this.GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a integer.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The value of the MAPI property as a integer.</returns>
        public int GetMapiPropertyInt32(string propIdentifier)
        {
            return (int)this.GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a byte array.
        /// </summary>
        /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
        /// <returns>The value of the MAPI property as a byte array.</returns>
        public byte[] GetMapiPropertyBytes(string propIdentifier)
        {
            return (byte[])this.GetMapiProperty(propIdentifier);
        }

        #endregion

        #region IDisposable Members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (!this.disposed)
            {
                //ensure only disposed once
                this.disposed = true;

                //call virtual disposing method to let sub classes clean up
                this.Disposing();

                //release COM storage object and suppress finalizer
                ReferenceManager.RemoveItem(this.storage);
                Marshal.ReleaseComObject(this.storage);
                GC.SuppressFinalize(this);
            }
        }

        /// <summary>
        /// Gives sub classes the chance to free resources during object disposal.
        /// </summary>
        protected virtual void Disposing() { }

        #endregion

    } //End OutlookStorage

} //End Namespace