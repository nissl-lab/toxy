namespace Toxy
{
    public interface ITextParser
    {
        string Parse();
        ParserContext Context { get; set; }
    }
}
