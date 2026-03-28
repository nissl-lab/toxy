namespace Toxy
{
	/// <summary>
	/// The <see cref="IEmailParser"/> is the interface for Parsers, which extracts an Email as structured Object (<see cref="ToxyEmail"/>).
	/// </summary>
	public interface IEmailParser
    {
		/// <summary>
		/// The Context of the <see cref="IEmailParser"/>.
		/// The Context contains the Path to the File or the <see cref="System.IO.Stream"/>.
		/// </summary>
		public ParserContext Context { get; set; }

		/// <summary>
		/// Parses the <see cref="Context"/> as structured Object (<see cref="ToxyEmail"/>).
		/// </summary>
		/// <returns>Returns the <see cref="ToxyEmail"/> of the File or <see cref="System.IO.Stream"/> in the <see cref="Context"/>.</returns>
		public ToxyEmail Parse();
    }
}
