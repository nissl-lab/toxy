#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenMcdf;
using Toxy.Ole;

namespace Toxy.Parsers;

public class OLE2MetadataParser : IMetadataParser
{
    public OLE2MetadataParser(ParserContext context)
    {
        Context = context;
    }

    public ParserContext Context { get; set; }

    public ToxyMetadata Parse()
    {
        Utility.ValidateContext(Context);
        Utility.ThrowIfProtected(Context);

        var metadata = new ToxyMetadata();
        using Stream stream = Utility.GetStream(Context);
        using var root = RootStorage.Open(stream, StorageModeFlags.LeaveOpen);

        if (OlePropertySet.TryRead(root, OlePropertySet.SummaryInformationStream,
                PropertySetReader.SummaryInfo, out var si))
            AddSummaryInformation(metadata, si);

        if (OlePropertySet.TryRead(root, OlePropertySet.DocumentSummaryInformationStream,
                PropertySetReader.DocSummaryInfo, out var dsi))
            AddDocumentSummaryInformation(metadata, dsi);

        return metadata;
    }

    private static void AddSummaryInformation(ToxyMetadata m, OlePropertySet si)
    {
        AddText(m, "Author", si.Text(4));
        AddText(m, "ApplicationName", si.Text(18));
        m.Add("ClassID", "{" + si.ClassId.ToString().ToUpperInvariant() + "}");
        AddNonZero(m, "CharCount", si.Int(16));
        AddNonZero(m, "ByteOrder", si.ByteOrder);
        AddDate(m, "CreateDateTime", si.Date(12));
        AddText(m, "LastAuthor", si.Text(8));
        AddText(m, "Keywords", si.Text(5));
        AddDate(m, "LastPrinted", si.Date(11));
        AddDate(m, "LastSaveDateTime", si.Date(13));
        AddNonZero(m, "PageCount", si.Int(14));
        AddNonZero(m, "WordCount", si.Int(15));
        AddText(m, "Comments", si.Text(6));
        if (si.FileTimeTicks(10) != 0) m.Add("EditTime", si.FileTimeTicks(10));
        AddText(m, "RevNumber", si.Text(9));
        AddNonZero(m, "Security", si.Int(19));
        AddText(m, "Subject", si.Text(3));
        AddText(m, "Title", si.Text(2));
        AddText(m, "Template", si.Text(7));
    }

    // The order is the priority of the former else-if chain.
    private static IEnumerable<(string Key, object Value)> DocumentSummaryCandidates(OlePropertySet d)
    {
        if (d.Int(4) > 0) yield return ("ByteCount", d.Int(4));
        if (d.Text(15) is { } company) yield return ("Company", company);
        if (d.Int(3) > 0) yield return ("Format", d.Int(3));
        if (d.Int(5) > 0) yield return ("LineCount", d.Int(5));
        if (d.Bool(16)) yield return ("LinksDirty", true);
        if (d.Text(14) is { } manager) yield return ("Manager", manager);
        if (d.Int(8) > 0) yield return ("NoteCount", d.Int(8));
        if (d.Bool(11)) yield return ("Scale", true);
        if (d.Int(9) > 0) yield return ("HiddenCount", d.Int(9));
        if (d.Int(10) > 0) yield return ("MMClipCount", d.Int(10));
        if (d.Int(6) > 0) yield return ("ParCount", d.Int(6));
    }

    private static void AddDocumentSummaryInformation(ToxyMetadata m, OlePropertySet dsi)
    {
        // Legacy behaviour: only the first match is added (like the former else-if chain).
        // Remove ".Take(1)" to add all available values.
        foreach (var (key, value) in DocumentSummaryCandidates(dsi).Take(1))
            m.Add(key, value);
    }

    private static void AddText(ToxyMetadata m, string key, string? value)
    {
        if (value != null) m.Add(key, value);
    }

    private static void AddNonZero(ToxyMetadata m, string key, int value)
    {
        if (value != 0) m.Add(key, value);
    }

    private static void AddDate(ToxyMetadata m, string key, DateTime? value)
    {
        if (value != null) m.Add(key, value.Value);
    }
}