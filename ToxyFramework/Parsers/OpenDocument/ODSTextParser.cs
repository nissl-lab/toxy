using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Toxy;
using Toxy.Parsers;

namespace ToxyFramework.Parsers.OpenDocument
{
	public class ODSTextParser : OpenDocumentTextParser
	{
		public ODSTextParser(ParserContext context) : base(context)
		{ }

		internal override string ParseText(XDocument document)
		{
			XNamespace tableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
			XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
			XName sheetName = tableNs + "table";
			XName sheetNameAttributeName = tableNs + "name";
			XName rowName = tableNs + "table-row";
			XName cellName = tableNs + "table-cell";
			XName pName = textNs + "p";
			XName reapetedName = tableNs + "number-columns-repeated";
			StringBuilder result = new StringBuilder();
			bool first = true;
			foreach (XElement sheet in document.Descendants(sheetName))
			{
				if (!first)
				{
					result.AppendLine();
				}

				string tableName = sheet.Attribute(sheetNameAttributeName)?.Value ?? "";
				int maxCols = -1;
				IEnumerable<IEnumerable<string>> tableData = GetTableData(sheet, rowName, cellName, pName, reapetedName);
				// get the length first
				// ODS Files often contains for example 1024 columns even if the User just uses 8
				foreach (IEnumerable<string> row in tableData)
				{
					int i = 0;
					int temp = -1;
					foreach (string col in row)
					{
						if (string.IsNullOrEmpty(col))
						{
							i++;
							continue;
						}

						temp = i;
						i++;
					}
					if (temp > maxCols)
					{
						maxCols = temp;
					}
				}
				if (maxCols < 0) continue;

				maxCols += 1;
				foreach (IEnumerable<string> row in tableData)
				{
#if NETSTANDARD2_1_OR_GREATER
					result.AppendJoin(',', row.Take(maxCols));
#else
					bool firstCol = true;
					foreach (string col in row.Take(maxCols))
					{
						if (!firstCol)
						{
							result.Append(',');
						}
						result.Append(col);
						firstCol = false;
					}
#endif
				}
				first = false;
			}
			return result.ToString();
		}

		private static IEnumerable<IEnumerable<string>> GetTableData(XElement sheet, XName rowName, XName cellName, XName pName, XName repeatedName)
		{
			foreach (XElement row in sheet.Descendants(rowName))
			{
				yield return GetRowData(row, cellName, pName, repeatedName);
			}
		}

		private static IEnumerable<string> GetRowData(XElement row, XName cellName, XName pName, XName repeatedName)
		{
			foreach (XElement cell in row.Elements(cellName))
			{
				string cellValue = string.Join(" ", cell.Descendants(pName).Select(p => p.Value));
				XAttribute? repeatAttr = cell.Attribute(repeatedName);
				int repeatCount = repeatAttr != null ? int.Parse(repeatAttr.Value) : 1;

				for (int i = 0; i < repeatCount; i++)
				{
					yield return cellValue;
				}
			}
		}
	}
}
