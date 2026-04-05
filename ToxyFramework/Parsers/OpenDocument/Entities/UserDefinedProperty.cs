using System.Xml.Serialization;

namespace Toxy.Parsers.OpenDocument.Entities
{
	[XmlRoot(ElementName = "user-defined", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
	public class UserDefinedProperty
	{
		[XmlAttribute(AttributeName = "name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string Name { get; set; }

		[XmlText]
		public string Value { get; set; }

		[XmlAttribute(AttributeName = "value-type", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string Type { get; set; }
	}
}
