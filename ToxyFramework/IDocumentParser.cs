namespace Toxy
{
    public interface IDocumentParser
    {
        ToxyDocument Parse();
        ParserContext Context { get; set; }
    }
}
