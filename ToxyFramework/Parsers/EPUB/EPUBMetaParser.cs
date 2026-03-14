using VersOne.Epub;

namespace Toxy.Parsers
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
			EpubBook book = EPUBHelper.GetEpubBook(Context);
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
