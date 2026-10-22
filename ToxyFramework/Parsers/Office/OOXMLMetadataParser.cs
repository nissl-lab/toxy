using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;

namespace Toxy.Parsers
{
	public class OOXMLMetadataParser : IMetadataParser
	{
		public OOXMLMetadataParser(ParserContext context)
		{
			Context = context;
		}

		public ParserContext Context { get; set; }

		public ToxyMetadata Parse()
		{
			Utility.ValidateContext(Context);

			ToxyMetadata metadata = new ToxyMetadata();
			Package pack = null;
			try
			{
				if (Context.IsStreamContext)
				{
					pack = Package.Open(Context.Stream, FileMode.Open, FileAccess.Read);
				}
				else
				{
					pack = Package.Open(Context.Path, FileMode.Open, FileAccess.Read, FileShare.Read);
				}

				AddCoreProperty(metadata, pack.PackageProperties);
				AddExtendedProperties(metadata, OoxmlExtendedProperties.Read(pack));
			}
			finally
			{
				pack?.Close();
			}
			return metadata;
		}

		/// <summary>
		/// Adds the first core property which is set. Only one of them is reported, this is the behavior this parser always had.
		/// </summary>
		private static void AddCoreProperty(ToxyMetadata metadata, PackageProperties props)
		{
			var candidates = new List<KeyValuePair<string, object>>
			{
				new KeyValuePair<string, object>("Title", props.Title),
				new KeyValuePair<string, object>("Identifier", props.Identifier),
				new KeyValuePair<string, object>("Keywords", props.Keywords),
				new KeyValuePair<string, object>("Revision", props.Revision),
				new KeyValuePair<string, object>("Subject", props.Subject),
				new KeyValuePair<string, object>("Modified", props.Modified),
				new KeyValuePair<string, object>("LastPrinted", props.LastPrinted),
				new KeyValuePair<string, object>("Created", props.Created),
				new KeyValuePair<string, object>("Creator", props.Creator),
				new KeyValuePair<string, object>("Description", props.Description),
			};
			foreach (KeyValuePair<string, object> candidate in candidates)
			{
				if (candidate.Value != null)
				{
					metadata.Add(candidate.Key, candidate.Value);
					return;
				}
			}
		}

		private static void AddExtendedProperties(ToxyMetadata metadata, IReadOnlyDictionary<string, string> props)
		{
			foreach (string name in new[] { "Application", "AppVersion", "Company", "Manager" })
			{
				if (props.TryGetValue(name, out string text))
				{
					metadata.Add(name, text);
				}
			}

			// counters are only reported if they are greater than zero
			foreach (string name in new[] { "Characters", "CharactersWithSpaces", "Lines", "Notes", "Pages", "Paragraphs", "Words", "TotalTime" })
			{
				if (props.TryGetValue(name, out string text)
					&& long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number)
					&& number > 0)
				{
					metadata.Add(name, number);
				}
			}
		}
	}
}
