#nullable enable
using System.Collections.Generic;
using System.Linq;

namespace Toxy.Ole;

public interface ITextExtractor
{
    IReadOnlyList<string> Extract(CompoundFile file);
}

/// <summary>Corel Presentations (.shw): the slide text is stored as ASCII in the stream "PerfectOffice_MAIN". The text is returned in fragments.</summary>
public sealed class CorelShowTextExtractor : ITextExtractor
{
    private static readonly string[] StreamNames =
    {
        "PerfectOffice_MAIN",
        "SlideShow/Main",
        OlePropertySet.SummaryInformationStream,
    };

    public IReadOnlyList<string> Extract(CompoundFile file) =>
        StreamNames
            .Select(file.TryReadStream)
            .Where(data => data is not null)
            .SelectMany(data => StringScanner.Ascii(data!))
            .ToList();
}

/// <summary>SolidWorks (.sldprt): only the metadata (property names and values) is readable, there is no model text.</summary>
public sealed class SolidWorksTextExtractor : ITextExtractor
{
    private static readonly string[] StreamNames =
    {
        OlePropertySet.SummaryInformationStream,
        OlePropertySet.DocumentSummaryInformationStream,
        "ISolidWorksInformation",
    };

    public IReadOnlyList<string> Extract(CompoundFile file) =>
        StreamNames
            .Select(file.TryReadStream)
            .Where(data => data is not null)
            .SelectMany(data => StringScanner.Ascii(data!).Concat(StringScanner.Utf16(data!)))
            .Distinct()
            .ToList();
}