module ToxyTypeProvider

open FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes

// The Toxy Type Provider entry point.
// Decorated with [<TypeProvider>] so the F# compiler discovers it at design time.
// Inherits TypeProviderForNamespaces from FSharp.TypeProviders.SDK and exposes
// a parameterised type ToxyDocument<SamplePath> whose Metadata members are
// generated at compile time by inspecting the sample file via MetadataSchema.

[<TypeProvider>]
type ToxyDocumentProvider(config: TypeProviderConfig) =
    inherit TypeProviderForNamespaces(config)

[<assembly: TypeProviderAssembly>]
do ()
