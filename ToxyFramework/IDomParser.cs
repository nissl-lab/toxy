namespace Toxy
{
    public interface IDomParser
    {
        ToxyDom Parse();
        ParserContext Context { get; set; }
    }
}
