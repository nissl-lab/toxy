using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Toxy.Parsers
{
	/// <summary>
	/// Reads the extended (application) properties of an Office Open XML package, i.e. the part <c>/docProps/app.xml</c>.
	/// </summary>
	internal static class OoxmlExtendedProperties
	{
		private const string RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

		/// <summary>
		/// Returns the text of the properties by their local name (e.g. <c>Application</c>, <c>Words</c>),
		/// or an empty dictionary if the package has no extended properties.
		/// </summary>
		public static IReadOnlyDictionary<string, string> Read(Package package)
		{
			PackageRelationship relationship = package.GetRelationshipsByType(RelationshipType).FirstOrDefault();
			if (relationship == null)
			{
				return new Dictionary<string, string>();
			}

			Uri partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
			if (!package.PartExists(partUri))
			{
				return new Dictionary<string, string>();
			}

			using (var stream = package.GetPart(partUri).GetStream())
			{
				XDocument document = XDocument.Load(stream);
				var properties = new Dictionary<string, string>(StringComparer.Ordinal);
				foreach (XElement element in document.Root.Elements())
				{
					properties[element.Name.LocalName] = element.Value;
				}
				return properties;
			}
		}
	}
}
