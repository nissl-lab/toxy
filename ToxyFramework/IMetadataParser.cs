namespace Toxy
{
    public interface IMetadataParser
    {
        ToxyMetadata Parse();
        ParserContext Context { get; set; }
    }
}
