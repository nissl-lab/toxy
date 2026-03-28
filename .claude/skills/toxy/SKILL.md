---
name: toxy
description: Use this skill whenever the user wants to extract text or data from documents using Toxy (the .NET text extraction library). Trigger when the user mentions reading, parsing, or extracting content from files like docx, xlsx, xls, pdf, csv, txt, epub, html, eml, vcf using C# or .NET. Also trigger when the user asks about Toxy NuGet package, ToxyDocument, ToxySpreadsheet, ToxyEmail, stream parsing, or any Toxy API usage. Use this skill for any Toxy 2.6 code generation, migration from older Toxy versions, or troubleshooting Toxy parsers.
---

# Toxy 2.6 Skill

Toxy is a .NET data/text extraction framework (similar to Apache Tika for Java). It supports cross-platform text and data extraction from many popular file formats. Always use **Toxy 2.6.0** (NuGet package `Toxy`, targeting `netstandard2.0` or `netstandard2.1`).

## Key Changes in 2.6

- Upgraded to `.NET Standard 2.1` support (in addition to 2.0)
- Added **stream-based parsing** (`ParserContext` now accepts `Stream` directly)
- Added **EPUB parser** (implements `IDocumentParser`)
- Removed unused `StreamReader` references (cleaner API)
- All parsers now live in `Toxy` namespace

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

The entry point for all parsing. Accepts either a **file path** or a **Stream** (new in 2.6):

```csharp
// From file path
var context = new ParserContext("path/to/file.docx");

// From stream (new in 2.6)
using var stream = File.OpenRead("path/to/file.docx");
var context = new ParserContext(stream, "docx"); // must supply format hint
```

### Parser Factory

Use `ParserFactory` to auto-detect format and return the correct parser:

```csharp
var parser = ParserFactory.CreateDocument(context);    // for document types
var parser = ParserFactory.CreateSpreadsheet(context); // for spreadsheet types
var parser = ParserFactory.CreateEmail(context);       // for email/contact types
```

---

## Toxy Object Types

| Object | Description | Formats |
|---|---|---|
| `ToxyDocument` | Paragraphs + metadata | docx, pdf, txt, epub, html, rtf, odt |
| `ToxySpreadsheet` | Rows/cells per sheet | xlsx, xls, csv, ods |
| `ToxyEmail` | Email fields | eml, msg |
| `ToxyBusinessCard` | Contact fields | vcf |
| `ToxyDom` | DOM tree | html, xml |
| `ToxyMetadata` | Key/value metadata | any file |

---

## Usage Patterns

### Extract Text from a Word Document

```csharp
using Toxy;

var context = new ParserContext("report.docx");
var parser = ParserFactory.CreateDocument(context);
ToxyDocument doc = parser.Parse();

foreach (var paragraph in doc.Paragraphs)
{
    Console.WriteLine(paragraph.Text);
}
```

### Extract Data from Excel

```csharp
using Toxy;

var context = new ParserContext("data.xlsx");
var parser = ParserFactory.CreateSpreadsheet(context);
ToxySpreadsheet sheet = parser.Parse();

foreach (var table in sheet.Tables)
{
    Console.WriteLine($"Sheet: {table.Name}");
    foreach (var row in table.Rows)
    {
        foreach (var cell in row.Cells)
        {
            Console.Write($"{cell.Value}\t");
        }
        Console.WriteLine();
    }
}
```

### Parse from a Stream (New in 2.6)

```csharp
using Toxy;

// Works with any stream source (MemoryStream, HttpResponseStream, etc.)
using var stream = File.OpenRead("document.pdf");
var context = new ParserContext(stream, "pdf");
var parser = ParserFactory.CreateDocument(context);
ToxyDocument doc = parser.Parse();
Console.WriteLine(doc.Paragraphs[0].Text);
```

### Parse PDF

```csharp
using Toxy;

var context = new ParserContext("file.pdf");
var parser = ParserFactory.CreateDocument(context);
ToxyDocument doc = parser.Parse();

foreach (var para in doc.Paragraphs)
    Console.WriteLine(para.Text);
```

### Parse EPUB (New in 2.6)

```csharp
using Toxy;

var context = new ParserContext("book.epub");
var parser = ParserFactory.CreateDocument(context);
ToxyDocument doc = parser.Parse();

foreach (var para in doc.Paragraphs)
    Console.WriteLine(para.Text);
```

### Parse Email

```csharp
using Toxy;

var context = new ParserContext("message.eml");
var parser = ParserFactory.CreateEmail(context);
ToxyEmail email = parser.Parse();

Console.WriteLine($"From: {email.From}");
Console.WriteLine($"Subject: {email.Subject}");
Console.WriteLine($"Body: {email.Body}");
```

### Parse Business Card (VCF)

```csharp
using Toxy;

var context = new ParserContext("contact.vcf");
var parser = ParserFactory.CreateEmail(context); // VCF uses email parser factory
ToxyBusinessCard card = (ToxyBusinessCard)parser.Parse();

Console.WriteLine(card.FullName);
Console.WriteLine(card.Email);
```

### Extract Metadata

```csharp
using Toxy;

var context = new ParserContext("file.pdf");
var parser = ParserFactory.CreateMetadata(context);
ToxyMetadata meta = parser.Parse();

foreach (var key in meta.Keys)
    Console.WriteLine($"{key}: {meta[key]}");
```

### Parse HTML as DOM

```csharp
using Toxy;

var context = new ParserContext("page.html");
var parser = ParserFactory.CreateDom(context);
ToxyDom dom = parser.Parse();

// Access DOM nodes
Console.WriteLine(dom.Root.InnerText);
```

---

## Supported Formats Summary

| Format | Extension(s) | Parser Type |
|---|---|---|
| Word (Open XML) | `.docx` | Document |
| Word (Legacy) | `.doc` | Document |
| PDF | `.pdf` | Document |
| Plain Text | `.txt` | Document |
| Rich Text | `.rtf` | Document |
| EPUB | `.epub` | Document (new in 2.6) |
| HTML | `.html`, `.htm` | Document / Dom |
| OpenDocument Text | `.odt` | Document |
| Excel (Open XML) | `.xlsx` | Spreadsheet |
| Excel (Legacy) | `.xls` | Spreadsheet |
| CSV | `.csv` | Spreadsheet |
| OpenDocument Sheet | `.ods` | Spreadsheet |
| Email | `.eml`, `.msg` | Email |
| Business Card | `.vcf` | Email (returns ToxyBusinessCard) |
| Any | `*` | Metadata |

---

## Tips & Best Practices

- **Auto-detection**: When using a file path, Toxy detects format from the extension automatically. When using a stream, always provide the format hint string (e.g., `"pdf"`, `"docx"`).
- **Error handling**: Wrap parse calls in try/catch — unsupported formats throw `NotSupportedException`.
- **Large files**: Use stream-based parsing to avoid loading entire files into memory.
- **Cross-platform**: Toxy targets netstandard2.0/2.1, so it works on Windows, Linux, and macOS.
- **No IFilter dependency**: Unlike old Windows-based approaches, Toxy does not require IFilter COM components.

For deeper reference on specific parsers and the class hierarchy, see `references/api.md`.
