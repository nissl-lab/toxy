module Tests

open System
open System.IO
open Xunit
open Toxy.TypeProvider.MetadataSchema

// Resolve path to a sample file from the Toxy.Test testdata folder.
// From the test output directory (bin/Debug/net8.0/), go up 4 levels to the
// repository root, then into the adjacent Toxy.Test testdata.
let private sampleDir =
    let assemblyDir = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
    Path.GetFullPath(Path.Combine(assemblyDir, "..", "..", "..", "..", "Toxy.Test", "testdata"))

let private docxSamplePath =
    Path.Combine(sampleDir, "ooxml", "MultipleCoreProperties.docx")

// ── null / empty path ────────────────────────────────────────────────────────

[<Fact>]
let ``extractSchema raises on null path`` () =
    Assert.Throws<Exception>(fun () -> extractSchema null |> ignore) |> ignore

[<Fact>]
let ``extractSchema raises on empty path`` () =
    Assert.Throws<Exception>(fun () -> extractSchema "" |> ignore) |> ignore

[<Fact>]
let ``extractSchema raises on whitespace path`` () =
    Assert.Throws<Exception>(fun () -> extractSchema "   " |> ignore) |> ignore

// ── non-existent file ────────────────────────────────────────────────────────

[<Fact>]
let ``extractSchema raises with meaningful message for missing file`` () =
    let ex = Assert.Throws<Exception>(fun () -> extractSchema "/tmp/no_such_file.docx" |> ignore)
    Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase)

// ── valid sample document ────────────────────────────────────────────────────

[<Fact>]
let ``extractSchema returns non-empty list for valid docx`` () =
    let schema = extractSchema docxSamplePath
    Assert.NotEmpty(schema)

[<Fact>]
let ``extractSchema returns non-null non-empty property names`` () =
    let schema = extractSchema docxSamplePath
    for prop in schema do
        Assert.False(String.IsNullOrEmpty(prop.Name), "Property name must not be null or empty")

[<Fact>]
let ``extractSchema returns non-null InferredType values`` () =
    let schema = extractSchema docxSamplePath
    for prop in schema do
        Assert.NotNull(prop.InferredType)
