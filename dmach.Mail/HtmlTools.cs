using System;
using System.IO;
using System.Text;

namespace InSolve.dmach.Mail
{
	/// <summary>
	/// Крутые вспомогательные методы
	/// </summary>
	internal abstract class HtmlTools
	{

		public static string GetHtmlFormatedText(string text)
		{
			if (text == null)
				throw new ArgumentNullException("text");

			StringBuilder buffer = new StringBuilder(
				text, 
				text.Length + Convert.ToInt32(text.Length * 0.2)
				);

			buffer.Replace("<", "&lt;");
			buffer.Replace(">", "&gt;");
			buffer.Replace("\n", "<br>");
			buffer.Replace("\r", string.Empty);
			
			return buffer.ToString();
		}

	}
}
