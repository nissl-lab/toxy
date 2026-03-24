namespace Toxy
{
	/// <summary>
	/// The <see cref="IMetadataParser"/> is the interface for Parsers, which extracts the <see cref="ToxyMetadata"/> of a File.
	/// </summary>
	public interface IMetadataParser
    {
		/// <summary>
		/// The Context of the <see cref="IEmailParser"/>.
		/// The Context contains the Path to the File or the <see cref="System.IO.Stream"/>.
		/// </summary>
		public ParserContext Context { get; set; }

		/// <summary>
		/// Parses the <see cref="ToxyMetadata"/> of the <see cref="Context"/>.
		/// </summary>
		/// <returns>Returns the <see cref="ToxyMetadata"/> of the File or <see cref="System.IO.Stream"/> in the <see cref="Context"/>.</returns>
		public ToxyMetadata Parse();
    }
}
