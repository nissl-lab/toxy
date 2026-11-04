#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;

namespace Toxy.Ole;

public sealed record MetadataEntry(string Name, object Value);

/// <summary>Generic extractor: returns all properties (including user defined ones) as a list of names and values.</summary>
public sealed class OleMetadataExtractor
{
    private static readonly Dictionary<uint, string> SummaryNames = new()
    {
        [2] = "Title", [3] = "Subject", [4] = "Author", [5] = "Keywords", [6] = "Comments",
        [7] = "Template", [8] = "LastSavedBy", [9] = "RevisionNumber", [10] = "TotalEditTime",
        [11] = "LastPrinted", [12] = "Created", [13] = "LastSaved", [14] = "Pages",
        [15] = "Words", [16] = "Characters", [18] = "Application", [19] = "Security",
    };

    private static readonly Dictionary<uint, string> DocSummaryNames = new()
    {
        [2] = "Category", [3] = "PresentationTarget", [4] = "Bytes", [5] = "Lines",
        [6] = "Paragraphs", [7] = "Slides", [8] = "Notes", [9] = "HiddenSlides",
        [10] = "MultimediaClips", [11] = "ScaleCrop", [12] = "HeadingPairs",
        [13] = "TitlesOfParts", [14] = "Manager", [15] = "Company", [16] = "LinksDirty",
        [17] = "CharactersWithSpaces", [19] = "SharedDoc", [22] = "HyperlinksChanged", [23] = "Version",
    };

    public IReadOnlyList<MetadataEntry> Extract(CompoundFile file)
    {
        var result = new List<MetadataEntry>();
        foreach (string streamName in new[]
                 {
                     OlePropertySet.SummaryInformationStream,
                     OlePropertySet.DocumentSummaryInformationStream,
                 })
        {
            var data = file.TryReadStream(streamName);
            if (data is null) continue;

            foreach (var section in PropertySetReader.Read(data))
            {
                var names = section.Format == PropertySetReader.SummaryInfo ? SummaryNames
                          : section.Format == PropertySetReader.DocSummaryInfo ? DocSummaryNames
                          : section.Names; // user defined properties (dictionary)

                // KeyValuePair cannot be deconstructed on .NET Standard 2.0
                foreach (var prop in section.Props.OrderBy(p => p.Key))
                {
                    string name = names.TryGetValue(prop.Key, out string? known) ? known : $"Property{prop.Key}";
                    var value = Normalize(name, prop.Value);
                    if (value is not null) result.Add(new MetadataEntry(name, value));
                }
            }
        }
        return result;
    }

    private static object? Normalize(string name, object? raw) => raw switch
    {
        FileTimeValue { Ticks: 0 } => null,
        FileTimeValue ft when name == "TotalEditTime" => TimeSpan.FromTicks(ft.Ticks),
        FileTimeValue ft => DateTime.FromFileTimeUtc(ft.Ticks),
        List<object?> list => string.Join("; ", list.Where(x => x is not null)),
        string { Length: 0 } => null,
        _ => raw,
    };
}