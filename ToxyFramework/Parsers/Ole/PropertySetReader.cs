#nullable enable
using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Ole;

/// <summary>Raw FILETIME value (100 ns ticks since 1601). Depending on the property it is a point in time or a duration.</summary>
public readonly struct FileTimeValue
{
    public FileTimeValue(long ticks) => Ticks = ticks;

    public long Ticks { get; }
}

/// <summary>Parser for OLE property sets (MS-OLEPS): SummaryInformation, DocumentSummaryInformation and user defined properties.</summary>
public static class PropertySetReader
{
    public static readonly Guid SummaryInfo = new("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
    public static readonly Guid DocSummaryInfo = new("D5CDD502-2E9C-101B-9397-08002B2CF9AE");
    public static readonly Guid UserDefined = new("D5CDD505-2E9C-101B-9397-08002B2CF9AE");

    public sealed class Section
    {
        public Section(Guid format, Dictionary<uint, object?> props, Dictionary<uint, string> names)
        {
            Format = format;
            Props = props;
            Names = names;
        }

        public Guid Format { get; }
        public Dictionary<uint, object?> Props { get; }
        public Dictionary<uint, string> Names { get; }
    }

    static PropertySetReader() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public static List<Section> Read(byte[] s)
    {
        var result = new List<Section>();
        if (s.Length < 48) return result;

        int count = BitConverter.ToInt32(s, 24);
        for (int i = 0; i < count && 28 + i * 20 + 20 <= s.Length; i++)
        {
            int e = 28 + i * 20;
            var format = new Guid(s.AsSpan(e, 16).ToArray());
            int offset = BitConverter.ToInt32(s, e + 16);
            if (offset > 0 && offset + 8 <= s.Length)
                result.Add(ReadSection(s, offset, format));
        }
        return result;
    }

    private static Section ReadSection(byte[] s, int offset, Guid format)
    {
        int count = BitConverter.ToInt32(s, offset + 4);
        var entries = new List<(uint Id, int Pos)>();
        for (int j = 0; j < count && offset + 8 + 8 * j + 8 <= s.Length; j++)
        {
            uint id = BitConverter.ToUInt32(s, offset + 8 + 8 * j);
            int pos = offset + BitConverter.ToInt32(s, offset + 12 + 8 * j);
            entries.Add((id, pos));
        }

        // Property 1 = code page (VT_I2, stored as signed value: -535 == 65001)
        int codePage = 1252;
        foreach (var (id, pos) in entries)
            if (id == 1 && pos >= 0 && pos + 6 <= s.Length)
                codePage = (ushort)BitConverter.ToInt16(s, pos + 4);
        var encoding = GetEncoding(codePage);

        var props = new Dictionary<uint, object?>();
        var names = new Dictionary<uint, string>();
        foreach (var (id, pos) in entries)
        {
            if (id == 1 || pos < 0 || pos + 4 > s.Length) continue;
            try
            {
                if (id == 0)
                {
                    ReadDictionary(s, pos, encoding, names);
                }
                else
                {
                    int p = pos;
                    props[id] = ReadValue(s, ref p, encoding);
                }
            }
            catch (ArgumentException)
            {
                // skip a corrupt property
            }
        }
        return new Section(format, props, names);
    }

    private static void ReadDictionary(byte[] s, int pos, Encoding encoding, Dictionary<uint, string> names)
    {
        int n = BitConverter.ToInt32(s, pos);
        int p = pos + 4;
        bool unicode = encoding.CodePage == 1200;
        for (int i = 0; i < n && p + 8 <= s.Length; i++)
        {
            uint id = BitConverter.ToUInt32(s, p);
            int length = BitConverter.ToInt32(s, p + 4);
            p += 8;
            int bytes = unicode ? length * 2 : length;
            names[id] = encoding.GetString(s, p, bytes).TrimEnd('\0');
            p += bytes;
            if (unicode) p = (p + 3) & ~3;
        }
    }

    private static object? ReadValue(byte[] s, ref int p, Encoding encoding)
    {
        ushort type = BitConverter.ToUInt16(s, p);
        p += 4;
        if ((type & 0x1000) == 0)
            return ReadScalar(s, ref p, type, encoding);

        // VT_VECTOR
        ushort elementType = (ushort)(type & 0x0FFF);
        int n = BitConverter.ToInt32(s, p);
        p += 4;
        var list = new List<object?>();
        for (int i = 0; i < n && p < s.Length; i++)
            list.Add(elementType == 12 ? ReadValue(s, ref p, encoding) : ReadScalar(s, ref p, elementType, encoding));
        return list;
    }

    private static object? ReadScalar(byte[] s, ref int p, ushort type, Encoding encoding)
    {
        object? value;
        switch (type)
        {
            case 2:
                value = BitConverter.ToInt16(s, p); p += 2; return value;
            case 18:
                value = BitConverter.ToUInt16(s, p); p += 2; return value;
            case 11:
                value = BitConverter.ToInt16(s, p) != 0; p += 2; return value;
            case 3 or 22:
                value = BitConverter.ToInt32(s, p); p += 4; return value;
            case 19 or 23:
                value = BitConverter.ToUInt32(s, p); p += 4; return value;
            case 20:
                value = BitConverter.ToInt64(s, p); p += 8; return value;
            case 21:
                value = BitConverter.ToUInt64(s, p); p += 8; return value;
            case 5:
                value = BitConverter.ToDouble(s, p); p += 8; return value;
            case 64:
                value = new FileTimeValue(BitConverter.ToInt64(s, p)); p += 8; return value;
            case 30: // VT_LPSTR (Codepage-kodiert)
            {
                int length = BitConverter.ToInt32(s, p);
                int bytes = encoding.CodePage == 1200 ? length * 2 : length;
                value = encoding.GetString(s, p + 4, bytes).TrimEnd('\0');
                p = (p + 4 + bytes + 3) & ~3;
                return value;
            }
            case 31: // VT_LPWSTR (UTF-16)
            {
                int length = BitConverter.ToInt32(s, p);
                value = Encoding.Unicode.GetString(s, p + 4, length * 2).TrimEnd('\0');
                p = (p + 4 + length * 2 + 3) & ~3;
                return value;
            }
            case 65 or 71: // blob / clipboard data (e.g. a thumbnail): skipped
                p += 4 + BitConverter.ToInt32(s, p);
                return null;
            default:
                return null;
        }
    }

    private static Encoding GetEncoding(int codePage)
    {
        try
        {
            return codePage == 1200 ? Encoding.Unicode : Encoding.GetEncoding(codePage);
        }
        catch (Exception e) when (e is ArgumentException or NotSupportedException)
        {
            return Encoding.GetEncoding(28591); // ISO-8859-1, Encoding.Latin1 does not exist in .NET Standard
        }
    }
}