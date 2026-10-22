---
name: toxy
description: Use this skill whenever the user wants to extract text or data from documents using Toxy (the .NET text extraction library). Trigger when the user mentions reading, parsing, or extracting content from files like docx, xlsx, xls, pptx, pdf, csv, txt, epub, html, eml, msg, vcf using C# or .NET. Also trigger when the user asks about Toxy NuGet package, ToxyDocument, ToxySpreadsheet, ToxyEmail, stream parsing, or any Toxy API usage. Use this skill for any Toxy 2.6 code generation, migration from older Toxy versions, or troubleshooting Toxy parsers.
---

# Toxy 2.6 Skill

Toxy is a .NET data/text extraction framework (similar to Apache Tika for Java). It supports cross-platform text and data extraction from many popular file formats. Always use **Toxy 2.6.0** (NuGet package `Toxy`, targeting `netstandard2.0`, `netstandard2.1`, `net8.0` and `net10.0`).

## Key Changes in 2.6

- Added `.NET Standard 2.1`, `net8.0` and `net10.0` targets (in addition to `netstandard2.0`)
- Added **stream-based parsing** (`ParserContext` accepts a `Stream`, the format is detected from the file signature)
- Added **EPUB parser** (text and metadata)
- Added **columnar spreadsheets** (`ParserFactory.CreateColumnarSpreadsheet`, csv/xls/xlsx)
- **NPOI removed**: `.xlsx` and `.docx` are read with the Open XML SDK (`DocumentFormat.OpenXml`), `.xls` with ExcelDataReader
- **Removed with NPOI**: the OLE2 metadata parser (`.ppt`, `.vsd`, `.pub`, `.shw`, `.sldprt` and metadata of `.xls`). Legacy Word `.doc` is not supported
- All parsers live in the `Toxy` namespace

---

## Installation

```xml
<!-- .csproj -->
<PackageReference Include="Toxy" Version="2.6.0" />
```

Or via CLI:
```bash
dotnet add package Toxy --version 2.6.0
```

---

## Core Concepts

### ParserContext

The entry point for all parsing. Accepts either a **file path** or a **Stream**:

```csharp
// From file path: the format is taken from the file extension
var context = new ParserContext("path/to/file.docx");

// From stream (new in 2.6): the format is detected from the file signature, there is no format hint
using var stream = File.OpenRead("path/to/file.docx");
var context = new ParserContext(stream);
```

Options are passed as string key/value pairs, `"1"`, `"true"` or `"on"` switch an option on:

```csharp
context.Properties.Add("HasHeader", "1");
```

### Parser Factory

`ParserFactory` picks the parser for the format of the context. Each method returns a different parser kind, not every kind exists for every format (see [Supported Formats](#supported-formats-summary)):

```csharp
ITextParser text                 = ParserFactory.CreateText(context);        // string
IDocumentParser document         = ParserFactory.CreateDocument(context);    // ToxyDocument
ISpreadsheetParser spreadsheet   = ParserFactory.CreateSpreadsheet(context); // ToxySpreadsheet
IEmailParser email               = ParserFactory.CreateEmail(context);       // ToxyEmail
VCardDocumentParser vcard        = ParserFactory.CreateVCard(context);       // ToxyBusinessCards
IDomParser dom                   = ParserFactory.CreateDom(context);         // ToxyDom
IMetadataParser metadata         = ParserFactory.CreateMetadata(context);    // ToxyMetadata
ISlideshowParser slideshow       = ParserFactory.CreateSlideshow(context);   // ToxySlideshow
IColumnarSpreadsheetParser cols  = ParserFactory.CreateColumnarSpreadsheet(context); // ToxyColumnarSpreadsheet
```

---

## Toxy Object Types

| Object | Description | Formats |
|---|---|---|
| `string` (text parser) | Plain text | almost all formats, see table below |
| `ToxyDocument` | Paragraphs, header, footer | docx, pdf |
| `ToxySpreadsheet` | Tables (sheets) with rows and cells | xlsx, xls, csv, pdf, html, eml, msg |
| `ToxyColumnarSpreadsheet` | Column oriented tables (Arrow/Parquet export) | xlsx, xls, csv |
| `ToxyEmail` | Email fields | eml, msg |
| `ToxyBusinessCards` | List of `ToxyBusinessCard` | vcf |
| `ToxySlideshow` | Slides with texts and notes | pptx |
| `ToxyDom` | DOM tree | html, xml |
| `ToxyMetadata` | Name/value properties | Office Open XML, OpenDocument, epub, eml, audio, images |

---

## Usage Patterns

### Extract Text (any supported format)

```csharp
using Toxy;

var context = new ParserContext("report.docx");
string text = ParserFactory.CreateText(context).Parse();
```

### Read a Word Document

```csharp
using Toxy;

var context = new ParserContext("report.docx");
ToxyDocument doc = ParserFactory.CreateDocument(context).Parse();

foreach (ToxyParagraph paragraph in doc.Paragraphs)
{
    Console.WriteLine($"[{paragraph.StyleID}] {paragraph.Text}");
}
```

### Extract Data from Excel

```csharp
using Toxy;

var context = new ParserContext("data.xlsx");
ToxySpreadsheet spreadsheet = ParserFactory.CreateSpreadsheet(context).Parse();

foreach (ToxyTable table in spreadsheet.Tables)
{
    Console.WriteLine($"Sheet: {table.Name}");
    foreach (ToxyRow row in table.Rows)
    {
        foreach (ToxyCell cell in row.Cells)
        {
            Console.Write($"[{cell.CellIndex}] {cell.Value}\t");
        }
        Console.WriteLine();
    }
}
```

Row and cell indexes are zero based. Empty cells are left out of `row.Cells` unless `FillBlankCells` is set, so use `cell.CellIndex` for the column. `table[i]` returns the row with `RowIndex == i` (or `null`).

#### Excel Options (`context.Properties`)

| Key | Default | Effect |
|---|---|---|
| `HasHeader` | off | The first row goes to `table.HeaderRows` instead of `table.Rows` |
| `FillBlankCells` | off | Empty cells are added with an empty value up to the last used column |
| `IncludeComments` | on | Cell comments are set on `ToxyCell.Comment` (text parser: appended as `Comment by <author>: ...`) |
| `ExtractSheetHeader` / `ExtractSheetFooter` | off | `table.PageHeader` / `table.PageFooter` as `left|center|right` |
| `ShowCalculatedResult` | off | Text parser: the calculated value instead of the formula text |
| `IncludeSheetNames` | on | Text parser: a line with the sheet name before each sheet |
| `IncludeHeaderFooter` | off | Text parser: page header and footer of each sheet |

Cell values are formatted with the number format of the cell (dates, decimals, currency), like Excel displays them.

#### Excel Limitations

- **`.xls`**: cell comments and formulas are not available (ExcelDataReader does not expose them). Cells which only carry formatting are not reported, so `FillBlankCells` fills up to the last column with content.
- **`.xlsx`**: for a shared formula only its first cell has the formula text, the other cells of the group return their calculated value.
- Encrypted workbooks throw an exception.

### Parse from a Stream (New in 2.6)

```csharp
using Toxy;

// Works with any stream source (MemoryStream, HttpResponseStream, etc.)
using var stream = File.OpenRead("document.pdf");
var context = new ParserContext(stream);
ToxyDocument doc = ParserFactory.CreateDocument(context).Parse();
Console.WriteLine(doc.Paragraphs[0].Text);
```

The stream must be readable from its start. Formats without a file signature (txt, csv, html, xml, vcf, eml, ...) cannot be detected from a stream; use a file path for them. An unknown signature throws `InvalidDataException`.

### Parse PDF

```csharp
using Toxy;

var context = new ParserContext("file.pdf");
ToxyDocument doc = ParserFactory.CreateDocument(context).Parse();

foreach (ToxyParagraph para in doc.Paragraphs)
    Console.WriteLine(para.Text);

// tables of the PDF
ToxySpreadsheet tables = ParserFactory.CreateSpreadsheet(context).Parse();
```

### Parse EPUB

EPUB supports the text and the metadata parser:

```csharp
using Toxy;

var context = new ParserContext("book.epub");
string text = ParserFactory.CreateText(context).Parse();
ToxyMetadata meta = ParserFactory.CreateMetadata(context).Parse();
```

### Parse Email

```csharp
using Toxy;

var context = new ParserContext("message.eml"); // or .msg
ToxyEmail email = ParserFactory.CreateEmail(context).Parse();

Console.WriteLine($"From: {email.From}");
Console.WriteLine($"To: {string.Join(", ", email.To)}");
Console.WriteLine($"Subject: {email.Subject}");
Console.WriteLine($"Body: {email.TextBody}");   // HtmlBody for the HTML part
```

### Parse Business Cards (VCF)

```csharp
using Toxy;

var context = new ParserContext("contacts.vcf");
ToxyBusinessCards cards = ParserFactory.CreateVCard(context).Parse();

foreach (ToxyBusinessCard card in cards.Cards)
{
    Console.WriteLine(card.Name.FullName);
    foreach (ToxyContact contact in card.Contacts)
        Console.WriteLine($"{contact.Name}: {contact.Value}");
}
```

### Parse a PowerPoint Slideshow

```csharp
using Toxy;

var context = new ParserContext("slides.pptx");
ToxySlideshow slideshow = ParserFactory.CreateSlideshow(context).Parse();

foreach (ToxySlide slide in slideshow.Slides)
    Console.WriteLine(string.Join(" ", slide.Texts));
```

### Extract Metadata

```csharp
using Toxy;

var context = new ParserContext("file.docx");
ToxyMetadata meta = ParserFactory.CreateMetadata(context).Parse();

foreach (ToxyProperty property in meta)
    Console.WriteLine($"{property.Name}: {property.Value}");

object title = meta.Get("Title")?.Value;
```

### Parse HTML as DOM

```csharp
using Toxy;

var context = new ParserContext("page.html");
ToxyDom dom = ParserFactory.CreateDom(context).Parse();

Console.WriteLine(dom.Root.InnerText);
ToxyNode title = dom.Root.SingleSelect("//title");
```

---

## Supported Formats Summary

| Format | Extension(s) | Parsers (`ParserFactory.Create…`) |
|---|---|---|
| Word (Open XML) | `.docx` | Text, Document, Metadata |
| Excel (Open XML) | `.xlsx` | Text, Spreadsheet, ColumnarSpreadsheet, Metadata |
| Excel (Legacy) | `.xls` | Text, Spreadsheet, ColumnarSpreadsheet |
| PowerPoint (Open XML) | `.pptx` | Text, Slideshow, Metadata |
| Visio / Publisher (Open XML) | `.vsdx`, `.pubx` | Metadata |
| PDF | `.pdf` | Text, Document, Spreadsheet |
| Plain Text | `.txt` | Text |
| CSV | `.csv` | Text, Spreadsheet, ColumnarSpreadsheet |
| XML | `.xml` | Text, Dom |
| HTML | `.html`, `.htm` | Text, Dom, Spreadsheet |
| Rich Text | `.rtf` | Text |
| EPUB | `.epub` | Text, Metadata |
| OpenDocument | `.odt`, `.odp`, `.ods` | Text, Metadata |
| Email | `.eml` | Email, Text, Spreadsheet, Metadata |
| Outlook Message | `.msg` | Email, Text, Spreadsheet |
| Business Card | `.vcf` | VCard, Text |
| Archives | `.zip`, `.gz`, `.tgz`, `.tar`, `.7z`, `.rar`, `.bz2`, `.lz`, `.xz`, `.boz` | Text (file list) |
| Audio | `.mp3`, `.ape`, `.wma`, `.flac`, `.aif` | Metadata |
| Images | `.jpg`, `.jpeg`, `.png`, `.gif`, `.tif`, `.tiff` | Metadata |

**Not supported**: legacy binary Office files other than `.xls` (`.doc`, `.ppt`, `.vsd`, `.pub`), including their metadata.

---

## Tips & Best Practices

- **Auto-detection**: With a file path the format comes from the extension, with a stream from the file signature. `ParserFactory.IsSupportedFileType(fileName)` checks an extension up front.
- **Error handling**: An unknown extension throws `NotSupportedException`; a parser kind the format does not offer (e.g. `CreateDocument` for `.xlsx`) and an undetectable stream throw `InvalidDataException`.
- **Large workbooks**: `.xlsx` sheets are read row by row and not loaded as a whole; `.xls` workbooks are read forward only.
- **Cross-platform**: Toxy targets netstandard2.0/2.1, net8.0 and net10.0, so it works on Windows, Linux, and macOS.
- **No IFilter dependency**: Unlike old Windows-based approaches, Toxy does not require IFilter COM components.

For deeper reference on specific parsers and the class hierarchy, see `references/api.md`.
