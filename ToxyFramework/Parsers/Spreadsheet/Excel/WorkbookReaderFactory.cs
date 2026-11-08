using System;
using System.IO;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Opens the right <see cref="IWorkbookReader"/> for a workbook by looking at its signature.
	/// </summary>
	internal static class WorkbookReaderFactory
	{
		private static readonly byte[] ZipSignature = { 0x50, 0x4B, 0x03, 0x04 };
		private static readonly byte[] Ole2Signature = { 0xD0, 0xCF, 0x11, 0xE0 };

		static WorkbookReaderFactory()
		{
			// legacy .xls files store text in code pages
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
		}

		/// <summary>
		/// Opens the workbook of <paramref name="context"/>. The returned reader disposes everything it opened itself,
		/// a stream passed in by the user stays open.
		/// </summary>
		public static IWorkbookReader Open(ParserContext context)
		{
			return Open(Utility.GetStream(context), ownsStream: !context.IsStreamContext);
		}

		/// <summary>
		/// Opens the workbook in <paramref name="stream"/>, which is read from its beginning.
		/// </summary>
		/// <param name="ownsStream">Whether the returned reader has to dispose <paramref name="stream"/>.</param>
		public static IWorkbookReader Open(Stream stream, bool ownsStream)
		{
			try
			{
				if (!stream.CanSeek)
				{
					var buffer = new MemoryStream();
					stream.CopyTo(buffer);
					if (ownsStream)
					{
						stream.Dispose();
					}
					stream = buffer;
					ownsStream = true;
				}
				stream.Position = 0;

				if (StartsWith(stream, ZipSignature))
				{
					return new OpenXmlWorkbookReader(stream, ownsStream);
				}
				if (StartsWith(stream, Ole2Signature))
				{
					return new ExcelDataReaderWorkbookReader(stream, ownsStream);
				}
				throw new InvalidDataException("The data is neither an .xlsx nor an .xls workbook.");
			}
			catch
			{
				if (ownsStream)
				{
					stream.Dispose();
				}
				throw;
			}
		}

		private static bool StartsWith(Stream stream, byte[] signature)
		{
			long position = stream.Position;
			try
			{
				var header = new byte[signature.Length];
				int read = 0;
				while (read < header.Length)
				{
					int count = stream.Read(header, read, header.Length - read);
					if (count == 0)
					{
						return false;
					}
					read += count;
				}
				for (int i = 0; i < signature.Length; i++)
				{
					if (header[i] != signature[i])
					{
						return false;
					}
				}
				return true;
			}
			finally
			{
				stream.Position = position;
			}
		}
	}
}
