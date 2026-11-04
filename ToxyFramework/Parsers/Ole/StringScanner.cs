#nullable enable
using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Ole;

/// <summary>Finds printable strings in binary data (ASCII or UTF-16LE).</summary>
public static class StringScanner
{
    public static IReadOnlyList<string> Ascii(ReadOnlySpan<byte> data, int minLength = 3)
    {
        var result = new List<string>();
        var sb = new StringBuilder();
        foreach (byte b in data)
        {
            if (b is >= 0x20 and < 0x7F)
            {
                sb.Append((char)b);
            }
            else
            {
                if (sb.Length >= minLength) result.Add(sb.ToString());
                sb.Clear();
            }
        }
        if (sb.Length >= minLength) result.Add(sb.ToString());
        return result;
    }

    public static IReadOnlyList<string> Utf16(ReadOnlySpan<byte> data, int minLength = 3)
    {
        var result = new List<string>();
        var sb = new StringBuilder();
        int i = 0;
        while (i + 1 < data.Length)
        {
            if (data[i] is >= 0x20 and < 0x7F && data[i + 1] == 0)
            {
                sb.Append((char)data[i]);
                i += 2;
            }
            else
            {
                if (sb.Length >= minLength) result.Add(sb.ToString());
                sb.Clear();
                i++;
            }
        }
        if (sb.Length >= minLength) result.Add(sb.ToString());
        return result;
    }
}