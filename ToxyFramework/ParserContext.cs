using FileSignatures;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy
{
	public class ParserContext
	{
		private Stream _filestream;
		public bool IsStreamContext { get => _filestream != null; }
		public Stream Stream { get => _filestream; }
		public string Path { get; set; }

		public string Extension
		{
			get
			{
				return Utility.GetFileExtention(this);
			}
		}
		public string MimeType
		{
			get
			{
				FileFormat fileformat = null;
				FileFormatInspector inspector = new FileFormatInspector();
				if (IsStreamContext)
				{
					Stream.Position = 0;
					fileformat = inspector.DetermineFileFormat(Stream);
				}
				else
				{
					using var stream = File.OpenRead(Path);
					fileformat = inspector.DetermineFileFormat(stream);
				}
				if (fileformat == null)
				{
					throw new InvalidDataException("File format could not be determined for the input stream");
				}
				return fileformat.MediaType;

			}
		}
		public Encoding Encoding { get; set; }

		public Dictionary<string, string> Properties { get; set; }

		public ParserContext(string path) : this(path, null)
		{ }
		public ParserContext(string path, Encoding encoding)
		{
			Path = path;
			Properties = new Dictionary<string, string>();
			Encoding = encoding;
		}
		public ParserContext(Stream stream) : this(null, Encoding.UTF8)
		{
			_filestream = stream;
		}
	}
}
