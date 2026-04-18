module MetadataSchema

// Design-time schema extraction module.
// Uses Toxy's ParserFactory and IMetadataParser to load a sample document,
// calls IMetadataParser.Parse() to obtain a ToxyMetadata instance, and
// iterates GetNames() to collect property names and infer their CLR types.
// The resulting schema is consumed by ToxyTypeProvider to generate
// strongly-typed F# provided types at compile/design time.
