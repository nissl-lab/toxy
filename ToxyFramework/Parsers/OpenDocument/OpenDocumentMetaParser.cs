using System.IO;
using System.IO.Compression;
using System.Xml.Serialization;
using ToxyFramework.Parsers.OpenDocument.Entities;

namespace Toxy.Parsers
{
	public sealed class OpenDocumentMetaParser : IMetadataParser
	{
		public ParserContext Context { get; set; }

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
			meta.Add("EditingCycles", body.EditingCycles);
			meta.Add("EditingDuration", body.EditingDuration);
			meta.Add("Generator", body.Generator);
			meta.Add("InitialCreator", body.InitialCreator);
			meta.Add("PrintedBy", body.PrintedBy);
			meta.Add("Creator", body.Creator);
			meta.Add("Description", body.Description);
			meta.Add("Language", body.Language);
			meta.Add("Subject", body.Subject);
			meta.Add("Title", body.Title);
			meta.Add("CreationDate", body.CreationDate);
			meta.Add("PrintDate", body.PrintDate);
			meta.Add("Date", body.Date);
			if (mimeType is not null)
			{
				meta.Add("MimeType", mimeType);
			}
			return meta;
#nullable restore
		}
	}
}
