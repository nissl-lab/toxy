using HtmlAgilityPack;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using VersOne.Epub;

namespace Toxy.Parsers
{
    public partial class EPUBParser : ITextParser
    {
        public EPUBParser(ParserContext context)
        {
            this.Context = context;
        }
        public virtual ParserContext Context
        {
            get; set;
        }
        public string Parse()
        {
            if (!File.Exists(Context.Path))
            {
                throw new FileNotFoundException("File " + Context.Path + " is not found");
            }

            StringBuilder result = new StringBuilder();
            HtmlDocument html = new HtmlDocument();
            EpubBook book = EpubReader.ReadBook(Context.Path);
            Regex whiteSpaceStart = new Regex("^\\s*", RegexOptions.Multiline);
            foreach (var chapter in book.ReadingOrder)
            {
                html.LoadHtml(chapter.Content);
                result.AppendLine(whiteSpaceStart.Replace(html.DocumentNode.InnerText, ""));
            }
            return result.ToString();
        }
    }
}
