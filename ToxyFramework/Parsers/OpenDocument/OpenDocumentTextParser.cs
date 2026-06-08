using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Toxy.Base;

namespace Toxy.Parsers
{
	public class OpenDocumentTextParser : BaseTextParser
	{

		public OpenDocumentTextParser(ParserContext context) : base(context)
		{ }

		internal override string ParseText(Stream stream)
		{
			// we need to let the stream open, which was passed by the user!
			using (ZipArchive archive = new ZipArchive(Context.Stream, ZipArchiveMode.Read, Context.IsStreamContext))
			{
				ZipArchiveEntry? contentEntry = archive.GetEntry("content.xml");
				if (contentEntry is null)
				{
					archive.Dispose();
					return "";
				}
				using (Stream contentStream = contentEntry.Open())
				{
					return ParseODFText(XDocument.Load(contentStream));
				}
			}
		}

		/// <summary>
		/// Parses the Text of the ODF File.
		/// The default implementation can extract the Text of the Basic Formats.
		/// </summary>
		/// <param name="stream">The Content <see cref="Stream"/> of the ODF File.</param>
		/// <returns>Returns the Text of the ODF File.</returns>
		protected virtual string ParseODFText(XDocument document)
		{
			XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
			// Leads to unexpected results at the moment
			// Need to parse the Text better in the future
			XName pName = textNs + "p";
			XName hName = textNs + "h";
			IEnumerable<XElement> paragraphs = document.Descendants().Where(x => x.Name == pName || x.Name == hName);
			return string.Join(Environment.NewLine, paragraphs.Select(p => p.Value));
		}
	}
}
