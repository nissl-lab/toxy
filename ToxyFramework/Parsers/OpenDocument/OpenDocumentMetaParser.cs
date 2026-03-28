using System.IO;
using System.IO.Compression;
using System.Xml.Serialization;
using ToxyFramework.Parsers.OpenDocument.Entities;

namespace Toxy.Parsers
{
	public sealed class OpenDocumentMetaParser : IMetadataParser
	{
		public ParserContext Context { get; set; }
		public OpenDocumentMetaParser(ParserContext context)
		{
			Context = context;
		}

		public ToxyMetadata Parse()
		{
#nullable enable
			ToxyMetadata meta = new ToxyMetadata();
			ZipArchive archive;
			if (!Context.IsStreamContext)
			{
				archive = ZipFile.OpenRead(Context.Path);
			}
			else
			{
				// we need to let the stream open, which was passed by the user!
				archive = new ZipArchive(Context.Stream, ZipArchiveMode.Read, true);
			}
			// Get metadata
			ZipArchiveEntry? entry = archive.GetEntry("meta.xml");
			if (entry is null)
			{
				archive.Dispose();
				return meta;
			}

			XmlSerializer serializer = new XmlSerializer(typeof(ODFMetaDocument));
			ODFMetaDocument? odfMetaDocument = null;
			using (Stream stream = entry.Open())
			{
				odfMetaDocument = (ODFMetaDocument?)serializer.Deserialize(stream);
			}
			if (odfMetaDocument is null || odfMetaDocument.Meta is null)
			{
				archive.Dispose();
				return meta;
			}

			string? mimeType = null;
			// contains the MimeType of the ODF.
			entry = archive.GetEntry("mimetype");
			if (entry is not null)
			{
				using (Stream stream = entry.Open())
				{
					using (StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8, true))
					{
						mimeType = reader.ReadToEnd();
					}
				}
			}

			ODFMetaBody body = odfMetaDocument.Meta;
			// Dublin Core
			meta.Add(nameof(body.Creator), body.Creator);
			meta.Add(nameof(body.Date), body.Date);
			meta.Add(nameof(body.Description), body.Description);
			meta.Add(nameof(body.Language), body.Language);
			meta.Add(nameof(body.Subject), body.Subject);
			meta.Add(nameof(body.Title), body.Title);
			meta.Add(nameof(body.Contributor), body.Contributor);
			meta.Add(nameof(body.Coverage), body.Coverage);
			meta.Add(nameof(body.Identifier), body.Identifier);
			meta.Add(nameof(body.Publisher), body.Publisher);
			meta.Add(nameof(body.Relation), body.Relation);
			meta.Add(nameof(body.Rights), body.Rights);
			meta.Add(nameof(body.Source), body.Source);
			meta.Add(nameof(body.Type), body.Type);
			// OpenDocument Format Meta
			meta.Add(nameof(body.CreationDate), body.CreationDate);
			meta.Add(nameof(body.EditingCycles), body.EditingCycles);
			meta.Add(nameof(body.EditingDuration), body.EditingDuration);
			meta.Add(nameof(body.Generator), body.Generator);
			meta.Add(nameof(body.InitialCreator), body.InitialCreator);
			meta.Add(nameof(body.Keywords), body.Keywords);
			meta.Add(nameof(body.PrintDate), body.PrintDate);
			meta.Add(nameof(body.PrintedBy), body.PrintedBy);
			if (mimeType is not null)
			{
				meta.Add("MimeType", mimeType);
			}
			return meta;
#nullable restore
		}
	}
}
