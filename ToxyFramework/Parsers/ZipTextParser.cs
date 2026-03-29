using SharpCompress.Archives;
using SharpCompress.Readers;
using System;
using System.IO;
using System.Text;
using Toxy.Base;
using Toxy.Helpers;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="ZipTextParser"/> is used to extract the Names of the Files & Directories inside compressed Files.
	/// </summary>
	public class ZipTextParser : BaseTextParser
	{
		/// <summary>
		/// Initializes the <see cref="ZipTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public ZipTextParser(ParserContext context) : base(context)
		{ }

		internal override string ParseText(out IDisposable disposable)
		{
			/* * NOTE ABOUT ENCRYPTION:
			 * In the ZIP file specification, encryption is applied at the entry level, not to the archive container itself.
			 * A single ZIP archive can contain a mix of encrypted and unencrypted files.
			 * Consequently, 'archive.IsEncrypted' may return false because the global directory doesn't always reflect the encryption status.
			 * To reliably determine if a password is required, you must check the 'IsEncrypted' property of the individual entries.
			*/

			Stream stream = Utility.GetStream(Context);
			// User Streams should not be closed!
			ReaderOptions options = new ReaderOptions() { LeaveStreamOpen = Context.IsStreamContext };
			IArchive archive = ArchiveFactory.OpenArchive(stream, options);
			// set here to make sure it will get disposed even if we throw an exception
			disposable = archive;
			if (archive.IsEncrypted)
			{
				ThrowHelper.ThrowEncrypted(Context.Path);
			}
			StringBuilder result = new StringBuilder();
			foreach (IArchiveEntry entry in archive.Entries)
			{
				if (entry.IsEncrypted)
				{
					ThrowHelper.ThrowEncrypted(entry.Key);
				}
				result.AppendLine(entry.Key);
			}
			return result.ToString();
		}
	}
}
