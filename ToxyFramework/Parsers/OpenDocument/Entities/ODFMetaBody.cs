using System;
using System.Collections.Generic;
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
		#region dublin core
		[XmlElement(ElementName = "creator", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Creator { get; set; } = "";
		
		[XmlElement(ElementName = "date", Namespace = "http://purl.org/dc/elements/1.1/")]
		public DateTimeOffset Date;
		
		[XmlElement(ElementName = "description", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Description { get; set; } = "";
		
		[XmlElement(ElementName = "language", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Language { get; set; } = "";
		
		[XmlElement(ElementName = "subject", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Subject { get; set; } = "";
		
		[XmlElement(ElementName = "title", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Title { get; set; } = "";


		[XmlElement(ElementName = "contributor", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Contributor { get; set; } = "";
		
		[XmlElement(ElementName = "coverage", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Coverage { get; set; } = "";
		
		[XmlElement(ElementName = "identifier", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Identifier { get; set; } = "";
		
		[XmlElement(ElementName = "publisher", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Publisher { get; set; } = "";
		
		[XmlElement(ElementName = "relation", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Relation { get; set; } = "";

		[XmlElement(ElementName = "rights", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Rights { get; set; } = "";

		[XmlElement(ElementName = "source", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Source { get; set; } = "";

		[XmlElement(ElementName = "type", Namespace = "http://purl.org/dc/elements/1.1/")]
		public string Type { get; set; } = "";
		#endregion

		#region Open Document Format Meta
		[XmlElement("creation-date", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public DateTimeOffset CreationDate;
		
		[XmlElement("editing-cycles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public uint EditingCycles;
		
		[XmlElement("editing-duration", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string EditingDuration { get; set; } = "";
		
		[XmlElement("generator", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string Generator { get; set; } = "";
		
		[XmlElement("initial-creator", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string InitialCreator { get; set; } = "";

		[XmlElement(ElementName = "keyword", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public List<string> Keywords { get; set; } = new List<string>();

		[XmlElement("print-date", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public DateTimeOffset PrintDate;
		
		[XmlElement("printed-by", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
		public string PrintedBy { get; set; } = "";
		#endregion

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
