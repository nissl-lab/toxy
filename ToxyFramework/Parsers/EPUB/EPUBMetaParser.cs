using HtmlAgilityPack;
using System.IO;
using System.Text;
using VersOne.Epub;

namespace Toxy.Parsers.EPUB
{
    public class EPUBMetaParser : IMetadataParser
    {
        public EPUBMetaParser(ParserContext context)
        {
            Context = context;
        }
        public virtual ParserContext Context { get; set; }
        public ToxyMetadata Parse()
        {
            if (!File.Exists(Context.Path))
            {
                throw new FileNotFoundException("File " + Context.Path + " is not found");
            }

            StringBuilder result = new StringBuilder();
            HtmlDocument html = new HtmlDocument();
            EpubBook book = EpubReader.ReadBook(Context.Path);
            ToxyMetadata meta = new ToxyMetadata();
            meta.Add(nameof(book.Author), book.Author);
            meta.Add(nameof(book.AuthorList), book.AuthorList);
            meta.Add(nameof(book.CoverImage), book.CoverImage);
            meta.Add(nameof(book.Description), book.Description);
            meta.Add(nameof(book.Title), book.Title);
            return meta;
        }
    }
}
