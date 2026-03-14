using NPOI.SS.Util;
using System.IO;
using VersOne.Epub;

namespace Toxy.Parsers
{
	/// <summary>
	/// Contains utility Methods for EPUB Files.
	/// </summary>
	internal static class EPUBHelper
	{
		/// <summary>
		/// Loads the <see cref="EpubBook"/> in <paramref name="context"/>
		/// </summary>
		/// <param name="context">Contains the Path to the EPUB File or the Stream.</param>
		/// <returns>Returns the <see cref="EpubBook"/> in <paramref name="context"/>.</returns>
		public static EpubBook GetEpubBook(ParserContext context)
		{
			Utility.ValidateContext(context);
			if (!context.IsStreamContext)
			{
				using (FileStream fs = new FileStream(context.Path, FileMode.Open, FileAccess.Read))
				{
					return EpubReader.ReadBook(fs);
				}
			}
			return EpubReader.ReadBook(context.Stream);
		}
	}
}
