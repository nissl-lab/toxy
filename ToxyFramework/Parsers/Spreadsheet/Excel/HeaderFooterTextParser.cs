using System.Text;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Splits a raw Excel header/footer code (e.g. <c>&amp;Cmy title&amp;RPage &amp;P</c>) into its left, center and right section.
	/// Formatting codes (fonts, sizes, colors, ...) and field codes (page number, date, ...) are dropped.
	/// </summary>
	internal static class HeaderFooterTextParser
	{
		public static HeaderFooterText Parse(string raw)
		{
			if (string.IsNullOrEmpty(raw))
			{
				return HeaderFooterText.Empty;
			}

			var left = new StringBuilder();
			var center = new StringBuilder();
			var right = new StringBuilder();
			StringBuilder current = center;

			for (int i = 0; i < raw.Length; i++)
			{
				char c = raw[i];
				if (c != '&' || i + 1 >= raw.Length)
				{
					current.Append(c);
					continue;
				}

				char code = raw[++i];
				switch (code)
				{
					case 'L': current = left; break;
					case 'C': current = center; break;
					case 'R': current = right; break;
					case '&': current.Append('&'); break;
					case '"':
						// font name: &"Arial,Bold"
						int close = raw.IndexOf('"', i + 1);
						i = close < 0 ? raw.Length : close;
						break;
					case 'K':
						// text color: &KRRGGBB
						i = System.Math.Min(i + 6, raw.Length - 1);
						break;
					default:
						if (char.IsDigit(code))
						{
							// font size: &12
							while (i + 1 < raw.Length && char.IsDigit(raw[i + 1]))
							{
								i++;
							}
						}
						// every other code (&P, &N, &D, &T, &F, &A, &B, &I, ...) is a field or style switch without text
						break;
				}
			}

			return new HeaderFooterText(left.ToString(), center.ToString(), right.ToString());
		}
	}
}
