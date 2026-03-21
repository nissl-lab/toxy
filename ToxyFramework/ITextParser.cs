namespace Toxy
{
	/// <summary>
	/// The <see cref="ITextParser"/> is the interface for Parsers, which extracts only the textual content of a File.
	/// </summary>
    public interface ITextParser
    {
		/// <summary>
		/// The Context of the <see cref="ITextParser"/>.
		/// The Context contains the Path to the File or the <see cref="System.IO.Stream"/>.
		/// </summary>
		public ParserContext Context { get; set; }

		/// <summary>
		/// Parses the Text of the <see cref="Context"/>.
		/// </summary>
		/// <returns>Returns the extracted Text of the File or <see cref="System.IO.Stream"/> in the <see cref="Context"/>.</returns>
		/// <remarks>
		/// This Method can throw different Exceptions those won't be catched.
		/// </remarks>
		public string Parse();
    }
}
