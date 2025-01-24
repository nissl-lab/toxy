namespace Toxy
{
    public interface IEmailParser
    {
        ToxyEmail Parse();
        ParserContext Context { get; set; }
    }
}
