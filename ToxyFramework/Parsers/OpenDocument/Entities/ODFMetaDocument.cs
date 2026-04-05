using System.Xml.Serialization;

namespace Toxy.Parsers.OpenDocument.Entities
{
	/// <summary>
	/// Represents the meta Document of an ODF (Open Document Format) File.
	/// <see href="https://docs.oasis-open.org/office/v1.1/OS/OpenDocument-v1.1-html/OpenDocument-v1.1.html#outline:2.2.Document_Metadata"/>
	/// </summary>
	/// Internal use only do not use.
	[XmlRoot(ElementName = "document-meta", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
	public sealed class ODFMetaDocument
	{
		/// <summary>
		/// Contains the actual Metadata of an ODF <see cref="ODFMetaBody"/>.
		/// </summary>
		[XmlElement("meta", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
		public ODFMetaBody Meta { get; set; } = new ODFMetaBody();
	}
}
