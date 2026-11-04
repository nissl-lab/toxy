#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Ole;

/// <summary>
/// Minimal, fully managed reader for OLE2 / Compound File Binary files (MS-CFB).
/// Reads all streams (including those stored in the mini stream) from a byte array.
/// </summary>
public sealed class CompoundFile
{
    private const uint MaxRegularSector = 0xFFFFFFFA;
    private const uint NoStream = 0xFFFFFFFF;

    private static readonly byte[] Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    private readonly byte[] _data;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly int _miniCutoff;
    private readonly uint[] _fat;
    private readonly uint[] _miniFat;
    private readonly byte[] _miniStream;
    private readonly List<StreamEntry> _streams = new();

    private readonly record struct StreamEntry(string Path, uint Start, long Size);

    public CompoundFile(byte[] data)
    {
        _data = data;
        if (data.Length < 512 || !data.AsSpan(0, 8).SequenceEqual(Signature))
            throw new InvalidDataException("The data is not an OLE2 compound file.");

        _sectorSize = 1 << BitConverter.ToUInt16(data, 30);
        _miniSectorSize = 1 << BitConverter.ToUInt16(data, 32);
        _miniCutoff = (int)BitConverter.ToUInt32(data, 56);

        uint firstDir = BitConverter.ToUInt32(data, 48);
        uint firstMiniFat = BitConverter.ToUInt32(data, 60);
        uint firstDifat = BitConverter.ToUInt32(data, 68);
        int numDifat = BitConverter.ToInt32(data, 72);

        // DIFAT -> list of the FAT sectors
        var fatSectors = new List<uint>();
        for (int i = 0; i < 109; i++)
            AddIfRegular(fatSectors, BitConverter.ToUInt32(data, 76 + 4 * i));

        uint difat = firstDifat;
        for (int n = 0; n < numDifat && difat < MaxRegularSector; n++)
        {
            int off = SectorOffset(difat);
            if (off < 0 || off + _sectorSize > data.Length) break;
            for (int i = 0; i < _sectorSize / 4 - 1; i++)
                AddIfRegular(fatSectors, BitConverter.ToUInt32(data, off + 4 * i));
            difat = BitConverter.ToUInt32(data, off + _sectorSize - 4);
        }

        _fat = ToUInts(fatSectors.SelectMany(SliceSector).ToArray());

        byte[] dir = ReadChain(firstDir);
        _miniFat = ToUInts(ReadChain(firstMiniFat));

        ParseDirectory(dir, out uint rootStart, out long rootSize);
        _miniStream = Trim(ReadChain(rootStart), rootSize);
    }

    public static CompoundFile Open(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return new CompoundFile(ms.ToArray());
    }

    /// <summary>All streams with their full path (e.g. "Contents/Config-0").</summary>
    public IEnumerable<(string Path, byte[] Data)> Streams() =>
        _streams.Select(e => (e.Path, ReadStream(e.Start, e.Size)));

    /// <summary>Reads a single stream; <c>null</c> if it does not exist.</summary>
    public byte[]? TryReadStream(string path)
    {
        foreach (var e in _streams)
            if (e.Path == path) return ReadStream(e.Start, e.Size);
        return null;
    }

    private void ParseDirectory(byte[] dir, out uint rootStart, out long rootSize)
    {
        int count = dir.Length / 128;
        if (count == 0) throw new InvalidDataException("The OLE2 compound file has an empty directory.");

        (byte Type, uint Left, uint Right, uint Child, uint Start, long Size, string Name) Entry(int i)
        {
            int o = i * 128;
            int nameBytes = Math.Max(0, Math.Min(BitConverter.ToUInt16(dir, o + 64) - 2, 62));
            return (dir[o + 66],
                    BitConverter.ToUInt32(dir, o + 68),
                    BitConverter.ToUInt32(dir, o + 72),
                    BitConverter.ToUInt32(dir, o + 76),
                    BitConverter.ToUInt32(dir, o + 116),
                    (long)BitConverter.ToUInt64(dir, o + 120),
                    Encoding.Unicode.GetString(dir, o, nameBytes));
        }

        var root = Entry(0);
        rootStart = root.Start;
        rootSize = root.Size;

        void Walk(uint id, string prefix, int depth)
        {
            if (id == NoStream || id >= count || depth > 1000) return;
            var e = Entry((int)id);
            Walk(e.Left, prefix, depth + 1);
            if (e.Type == 2) _streams.Add(new StreamEntry(prefix + e.Name, e.Start, e.Size));
            else if (e.Type == 1) Walk(e.Child, prefix + e.Name + "/", depth + 1);
            Walk(e.Right, prefix, depth + 1);
        }

        Walk(root.Child, "", 0);
    }

    private byte[] ReadStream(uint start, long size)
    {
        if (size >= _miniCutoff)
            return Trim(ReadChain(start), size);

        var ms = new MemoryStream();
        uint id = start;
        int guard = 0;
        while (id < MaxRegularSector && id < _miniFat.Length && guard++ <= _miniFat.Length)
        {
            int o = (int)id * _miniSectorSize;
            if (o >= _miniStream.Length) break;
            ms.Write(_miniStream, o, Math.Min(_miniSectorSize, _miniStream.Length - o));
            id = _miniFat[id];
        }
        return Trim(ms.ToArray(), size);
    }

    private byte[] ReadChain(uint start)
    {
        var ms = new MemoryStream();
        uint id = start;
        int guard = 0;
        while (id < MaxRegularSector && id < _fat.Length && guard++ <= _fat.Length)
        {
            int o = SectorOffset(id);
            if (o < 0 || o >= _data.Length) break;
            ms.Write(_data, o, Math.Min(_sectorSize, _data.Length - o));
            id = _fat[id];
        }
        return ms.ToArray();
    }

    private int SectorOffset(uint id)
    {
        long o = (id + 1L) * _sectorSize;
        return o > int.MaxValue ? -1 : (int)o;
    }

    private byte[] SliceSector(uint id)
    {
        int o = SectorOffset(id);
        if (o < 0 || o >= _data.Length) return Array.Empty<byte>();
        return _data.AsSpan(o, Math.Min(_sectorSize, _data.Length - o)).ToArray();
    }

    private static byte[] Trim(byte[] bytes, long size) =>
        size >= 0 && size < bytes.Length ? bytes.AsSpan(0, (int)size).ToArray() : bytes;

    private static void AddIfRegular(List<uint> list, uint value)
    {
        if (value < MaxRegularSector) list.Add(value);
    }

    private static uint[] ToUInts(byte[] bytes)
    {
        var result = new uint[bytes.Length / 4];
        Buffer.BlockCopy(bytes, 0, result, 0, result.Length * 4);
        return result;
    }
}