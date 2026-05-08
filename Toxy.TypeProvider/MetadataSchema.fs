module Toxy.TypeProvider.MetadataSchema

open System
open Toxy

/// Represents a single inferred property from document metadata
type SchemaProperty = {
    Name: string
    InferredType: Type
}

/// Infers a CLR Type from a metadata value object
let private inferType (value: obj) : Type =
    match value with
    | :? DateTime -> typeof<DateTime>
    | :? int      -> typeof<int>
    | :? int64    -> typeof<int64>
    | :? float    -> typeof<float>
    | :? bool     -> typeof<bool>
    | _           -> typeof<string>

/// Extracts a schema (list of SchemaProperty) from a sample document at the given file path.
/// Uses Toxy's ParserFactory and IMetadataParser at design/compile time.
let extractSchema (samplePath: string) : SchemaProperty list =
    if String.IsNullOrWhiteSpace(samplePath) then
        failwith "samplePath must not be null, empty, or whitespace"
    if not (System.IO.File.Exists(samplePath)) then
        failwithf "Sample file not found: %s" samplePath
    let ctx = ParserContext(samplePath)
    let parser = ParserFactory.CreateMetadata(ctx)
    if isNull parser then
        failwithf "No metadata parser available for file: %s" samplePath
    let metadata = parser.Parse()
    if isNull metadata then []
    else
        metadata.GetNames()
        |> Array.toList
        |> List.map (fun name ->
            let prop = metadata.Get(name)
            let inferredType = if isNull prop.Value then typeof<string> else inferType prop.Value
            { Name = name; InferredType = inferredType })
