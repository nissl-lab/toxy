# Toxy 2.6 API Reference

## Namespace
All types are under the `Toxy` namespace.

---

## ParserContext

The unified context object passed to all parser factories.

```csharp
// Constructor overloads
public ParserContext(string filePath)
public ParserContext(Stream stream, string formatHint)  // new in 2.6
```

**Properties:**
- `string FilePath` — path to the file (null if stream-based)
- `Stream Stream` — underlying stream (null if path-based)
- `string FormatHint` — format extension hint used for stream parsing

---

## ParserFactory (static)

```csharp
// Returns IDocumentParser
ParserFactory.CreateDocument(ParserContext context)

// Returns ISpreadsheetParser
ParserFactory.CreateSpreadsheet(ParserContext context)

// Returns IEmailParser (also used for VCF)
ParserFactory.CreateEmail(ParserContext context)

// Returns IMetadataParser
ParserFactory.CreateMetadata(ParserContext context)

// Returns IDomParser
ParserFactory.CreateDom(ParserContext context)
```

All methods throw `NotSupportedException` if the format is not recognized or not supported for the given parser type.

---

## ToxyDocument

Returned by `IDocumentParser.Parse()`.

```csharp
public class ToxyDocument
{
    public List<ToxyParagraph> Paragraphs { get; }
    public ToxyMetadata Metadata { get; }
    public string PageCount { get; }    // for PDFs
}

public class ToxyParagraph
{
    public string Text { get; }
    public int Level { get; }           // heading level (0 = body text)
    public bool IsHeading { get; }
}
```

---

## ToxySpreadsheet

Returned by `ISpreadsheetParser.Parse()`.

```csharp
public class ToxySpreadsheet
{
    public List<ToxyTable> Tables { get; }  // one per worksheet
}

public class ToxyTable
{
    public string Name { get; }             // sheet name
    public List<ToxyRow> Rows { get; }
}

public class ToxyRow
{
    public int RowIndex { get; }
    public List<ToxyCell> Cells { get; }
}

public class ToxyCell
{
    public int ColumnIndex { get; }
    public object Value { get; }            // string, double, DateTime, bool, null
    public string Formula { get; }          // null if not a formula cell
}
```

---

## ToxyEmail

Returned by `IEmailParser.Parse()` for `.eml` / `.msg` files.

```csharp
public class ToxyEmail
{
    public string From { get; }
    public List<string> To { get; }
    public List<string> Cc { get; }
    public string Subject { get; }
    public string Body { get; }
    public bool IsHtml { get; }
    public DateTime Date { get; }
    public List<ToxyAttachment> Attachments { get; }
}

public class ToxyAttachment
{
    public string Name { get; }
    public byte[] Data { get; }
}
```

---

## ToxyBusinessCard

Returned by `IEmailParser.Parse()` for `.vcf` files (cast from result).

```csharp
public class ToxyBusinessCard
{
    public string FullName { get; }
    public string FirstName { get; }
    public string LastName { get; }
    public string Email { get; }
    public string Phone { get; }
    public string Organization { get; }
    public string Title { get; }
    public string Address { get; }
    public string Url { get; }
}
```

Usage:
```csharp
var parser = ParserFactory.CreateEmail(context);
var card = (ToxyBusinessCard)parser.Parse();
```

---

## ToxyDom

Returned by `IDomParser.Parse()`.

```csharp
public class ToxyDom
{
    public ToxyNode Root { get; }
}

public class ToxyNode
{
    public string Name { get; }
    public string InnerText { get; }
    public string OuterHtml { get; }
    public List<ToxyNode> Children { get; }
    public Dictionary<string, string> Attributes { get; }
}
```

---

## ToxyMetadata

Returned by `IMetadataParser.Parse()`, also embedded in `ToxyDocument`.

```csharp
public class ToxyMetadata : IEnumerable<KeyValuePair<string, object>>
{
    public IEnumerable<string> Keys { get; }
    public object this[string key] { get; }
    public bool Contains(string key);
}
```

---

## Parser Interfaces

```csharp
public interface IDocumentParser
{
    ToxyDocument Parse();
}

public interface ISpreadsheetParser
{
    ToxySpreadsheet Parse();
}

public interface IEmailParser
{
    // Returns ToxyEmail or ToxyBusinessCard depending on format
    object Parse();
}

public interface IMetadataParser
{
    ToxyMetadata Parse();
}

public interface IDomParser
{
    ToxyDom Parse();
}
```

---

## Common Patterns

### Check if Cell Has Value
```csharp
foreach (var cell in row.Cells)
{
    if (cell.Value != null)
        Console.WriteLine($"[{cell.ColumnIndex}] = {cell.Value}");
}
```

### Filter Headings Only
```csharp
var headings = doc.Paragraphs.Where(p => p.IsHeading);
```

### Stream from HttpClient (ASP.NET Core example)
```csharp
var response = await httpClient.GetAsync(url);
var stream = await response.Content.ReadAsStreamAsync();
var context = new ParserContext(stream, "pdf");
var parser = ParserFactory.CreateDocument(context);
var doc = parser.Parse();
```

### Batch File Processing
```csharp
var files = Directory.GetFiles("./docs", "*.*");
foreach (var file in files)
{
    try
    {
        var ctx = new ParserContext(file);
        var ext = Path.GetExtension(file).TrimStart('.').ToLower();
        
        if (ext is "xlsx" or "xls" or "csv")
        {
            var sheet = ParserFactory.CreateSpreadsheet(ctx).Parse();
            // process sheet...
        }
        else
        {
            var doc = ParserFactory.CreateDocument(ctx).Parse();
            // process doc...
        }
    }
    catch (NotSupportedException)
    {
        Console.WriteLine($"Unsupported format: {file}");
    }
}
```
