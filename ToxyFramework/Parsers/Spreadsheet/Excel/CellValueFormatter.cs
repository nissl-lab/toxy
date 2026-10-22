using ExcelNumberFormat;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Converts raw cell values to the text Excel would display, so that every
	/// <see cref="IWorkbookReader"/> produces identical output for the same data.
	/// </summary>
	internal static class CellValueFormatter
	{
		private const string General = "General";

		// Format codes of the built-in number formats (ECMA-376, 18.8.30), in case a library only reports the id.
		private static readonly Dictionary<int, string> BuiltInFormats = new Dictionary<int, string>
		{
			{ 1, "0" }, { 2, "0.00" }, { 3, "#,##0" }, { 4, "#,##0.00" },
			{ 9, "0%" }, { 10, "0.00%" }, { 11, "0.00E+00" }, { 12, "# ?/?" }, { 13, "# ??/??" },
			{ 14, "mm-dd-yy" }, { 15, "d-mmm-yy" }, { 16, "d-mmm" }, { 17, "mmm-yy" },
			{ 18, "h:mm AM/PM" }, { 19, "h:mm:ss AM/PM" }, { 20, "h:mm" }, { 21, "h:mm:ss" }, { 22, "m/d/yy h:mm" },
			{ 37, "#,##0 ;(#,##0)" }, { 38, "#,##0 ;[Red](#,##0)" }, { 39, "#,##0.00;(#,##0.00)" }, { 40, "#,##0.00;[Red](#,##0.00)" },
			{ 45, "mm:ss" }, { 46, "[h]:mm:ss" }, { 47, "mmss.0" }, { 48, "##0.0E+0" }, { 49, "@" },
		};

		private static readonly ConcurrentDictionary<string, NumberFormat> Cache = new ConcurrentDictionary<string, NumberFormat>(StringComparer.Ordinal);

		/// <summary>Resolves the format code of a built-in number format id, or <c>null</c> if the id is not a known built-in format.</summary>
		public static string GetBuiltInFormat(int numberFormatId)
		{
			BuiltInFormats.TryGetValue(numberFormatId, out string format);
			return format;
		}

		/// <summary>
		/// Formats <paramref name="value"/> with <paramref name="numberFormat"/>.
		/// </summary>
		/// <param name="value">A <c>null</c>, <see cref="string"/>, <see cref="bool"/>, number, <see cref="DateTime"/> or <see cref="TimeSpan"/>.</param>
		/// <param name="numberFormat">The Excel format code. <c>null</c> or empty means "General".</param>
		/// <param name="isDate1904">Whether a number formatted as date counts the days from 1904 instead of 1900.</param>
		public static string Format(object value, string numberFormat, bool isDate1904 = false)
		{
			switch (value)
			{
				case null:
					return string.Empty;
				case string text:
					return text;
				case bool flag:
					return flag ? "TRUE" : "FALSE";
				case TimeSpan span:
					return FormatNumber(span.TotalDays, numberFormat, isDate1904);
				case DateTime _:
					return FormatNumber(value, numberFormat, isDate1904);
				case IConvertible convertible when IsNumeric(value):
					return FormatNumber(convertible.ToDouble(CultureInfo.InvariantCulture), numberFormat, isDate1904);
				default:
					return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
			}
		}

		private static bool IsNumeric(object value)
		{
			switch (Type.GetTypeCode(value.GetType()))
			{
				case TypeCode.Byte:
				case TypeCode.SByte:
				case TypeCode.Int16:
				case TypeCode.UInt16:
				case TypeCode.Int32:
				case TypeCode.UInt32:
				case TypeCode.Int64:
				case TypeCode.UInt64:
				case TypeCode.Single:
				case TypeCode.Double:
				case TypeCode.Decimal:
					return true;
				default:
					return false;
			}
		}

		private static string FormatNumber(object value, string numberFormat, bool isDate1904)
		{
			string code = string.IsNullOrEmpty(numberFormat) ? General : numberFormat;
			NumberFormat format = Cache.GetOrAdd(code, c => new NumberFormat(c));
			if (!format.IsValid)
			{
				format = Cache.GetOrAdd(General, c => new NumberFormat(c));
			}
			return format.Format(value, CultureInfo.InvariantCulture, isDate1904);
		}
	}
}
