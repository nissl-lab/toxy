using System;
using System.Xml.Serialization;

namespace ToxyFramework.Parsers.OpenDocument.Entities
{
	/// <summary>
	/// Represents the Metadata of an ODF (Open Document Format) File.
	/// <see href="https://docs.oasis-open.org/office/v1.1/OS/OpenDocument-v1.1-html/OpenDocument-v1.1.html#outline:3.Metadata_Elements"/>
	/// </summary>
	/// Internal use only do not use.
	public sealed class ODFMetaBody
	{
		[XmlElement(ElementName = "creator", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Creator { get; set; } = "";
		[XmlElement(ElementName = "date", Namespace = "http://purl.org/dc/elements/1.1/")]
		public DateTime Date;
		[XmlElement(ElementName = "description", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Description { get; set; } = "";
		[XmlElement(ElementName = "language", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Language { get; set; } = "";
		[XmlElement(ElementName = "subject", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Subject { get; set; } = "";
		[XmlElement(ElementName = "title", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Title { get; set; } = "";

		[XmlElement("creation-date", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public DateTime CreationDate;
		[XmlElement("editing-cycles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public uint EditingCycles;
		[XmlElement("editing-duration", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string EditingDuration { get; set; } = "";
		[XmlElement("generator", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string Generator { get; set; } = "";
		[XmlElement("initial-creator", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string InitialCreator { get; set; } = "";
		[XmlElement("keyword", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string Keyword { get; set; } = "";
		[XmlElement("print-date", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public DateTime PrintDate;
		[XmlElement("printed-by", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string PrintedBy { get; set; } = "";

		#region Missing
		/*
         * template
         * auto-reload
         * hyperlink-behaviour
         * document-statistic
         * meta:user-defined
        */
		#endregion
	}
}
