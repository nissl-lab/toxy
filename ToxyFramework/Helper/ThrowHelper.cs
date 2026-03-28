using System;
using System.Diagnostics.CodeAnalysis;

namespace Toxy.Helper
{
	/// <summary>
	/// Contains Methods, which throws an <see cref="Exception"/>.
	/// </summary>
	internal static class ThrowHelper
	{
#nullable enable
		/// <summary>
		/// Throws an <see cref="InvalidOperationException"/> indicating that the specified file is encrypted.
		/// If the <paramref name="filePath"/> is not specified a more general exception will be thrown without additional Data.
		/// </summary>
		/// <param name="filePath">The full path or name of the file that is encrypted.</param>
		/// <exception cref="InvalidOperationException">Always thrown to signal that encrypted files are not supported.</exception>
#if NETCOREAPP2_1_OR_GREATER
		[DoesNotReturn]
		#endif
		public static void ThrowEncrypted(string? filePath = null)
		{
			if (string.IsNullOrWhiteSpace(filePath))
			{
				throw new InvalidOperationException($"file {filePath} is encrypted");
			}
			throw new InvalidOperationException($"stream or file is encrypted");
		}
#nullable restore
	}
}
