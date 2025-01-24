namespace Toxy
{
    public interface ISlideshowParser
    {
        ToxySlideshow Parse();
        ToxySlide Parse(int slideIndex);
        ParserContext Context { get; set; }
    }
}
