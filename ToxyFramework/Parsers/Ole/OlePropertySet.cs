#nullable enable
using System;
using System.IO;
using OpenMcdf;

namespace Toxy.Ole;

/// <summary>Typed access to the first section of a property set (replaces NPOI SummaryInformation/DocumentSummaryInformation).</summary>
public sealed class OlePropertySet
{
    public const string SummaryInformationStream = "\u0005SummaryInformation";
    public const string DocumentSummaryInformationStream = "\u0005DocumentSummaryInformation";

    private readonly PropertySetReader.Section _section;

    private OlePropertySet(Guid classId, int byteOrder, PropertySetReader.Section section)
    {
        ClassId = classId;
        ByteOrder = byteOrder;
        _section = section;
    }

    public Guid ClassId { get; }
    public int ByteOrder { get; }

    public static bool TryRead(RootStorage root, string streamName, Guid expectedFormat, out OlePropertySet set)
    {
        set = null!;
        if (!root.TryOpenStream(streamName, out CfbStream? stream) || stream is null) return false;

        byte[] data;
        using (stream)
        using (var buffer = new MemoryStream())
        {
            stream.CopyTo(buffer);
            data = buffer.ToArray();
        }
        if (data.Length < 48) return false;

        var sections = PropertySetReader.Read(data);
        if (sections.Count == 0 || sections[0].Format != expectedFormat) return false;

        set = new OlePropertySet(new Guid(data.AsSpan(8, 16).ToArray()), BitConverter.ToUInt16(data, 0), sections[0]);
        return true;
    }

    public string? Text(uint id) => Get(id) as string;

    public int Int(uint id) => Get(id) switch
    {
        int i => i,
        short s => s,
        uint u => (int)u,
        ushort us => us,
        _ => 0,
    };

    public bool Bool(uint id) => Get(id) is true;

    public long FileTimeTicks(uint id) => Get(id) switch
    {
        FileTimeValue ft => ft.Ticks,
        long l => l,
        _ => 0,
    };

    public DateTime? Date(uint id)
    {
        long ticks = FileTimeTicks(id);
        if (ticks <= 0) return null;
        try { return DateTime.FromFileTimeUtc(ticks); }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    private object? Get(uint id) => _section.Props.TryGetValue(id, out object? value) ? value : null;
}