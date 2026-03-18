using DocumentFormat.OpenXml.InkML;
using FileSignatures;
using PasswordProtectedChecker;
using System;
using System.IO;

namespace Toxy
{
    public static class Utility
    {
        /// <summary>
        /// Used to check if a File/Stream is protected.
        /// </summary>
        private static readonly Checker _checker = new Checker();

        public static bool IsTrue(string sValue)
        {
            return (sValue.Equals("1", StringComparison.OrdinalIgnoreCase)) || (sValue.Equals("on", StringComparison.OrdinalIgnoreCase)) || (sValue.Equals("true", StringComparison.OrdinalIgnoreCase));
        }

        public static void ValidateContext(ParserContext context)
        {
            if (!context.IsStreamContext && !File.Exists(context.Path))
                throw new FileNotFoundException("File " + context.Path + " is not found");

            if (context.IsStreamContext && context.Stream == null)
                throw new InvalidOperationException("Context.Stream is null");
        }
        public static Stream GetStream(ParserContext context)
        {
            if (context.IsStreamContext)
                return context.Stream;
            else
                return File.OpenRead(context.Path);
        }

        public static string GetFileExtention(ParserContext context)
        {
            string ext = null;
            if (context.IsStreamContext)
            {
                FileFormatInspector inspector = new FileFormatInspector();
                context.Stream.Position = 0;
                var fileformat = inspector.DetermineFileFormat(context.Stream);
                if (fileformat == null)
                    throw new InvalidDataException("File format could not be determined for the input stream");
                ext = '.' + fileformat.Extension;
            }
            else
            {
                FileInfo fi = new FileInfo(context.Path);
                ext = fi.Extension;
            }
            return ext;
        }

        public static void ThrowIfProtected(ParserContext context)
        {
            Result result;
			if (!context.IsStreamContext)
			{
                result = _checker.IsFileProtected(context.Path);
			}
			else
			{
                result = _checker.IsStreamProtected(context.Stream, GetFileExtention(context));
			}

			if (result.Protected)
			{
                // IDK what should be thrown in case of Streams.
				throw new InvalidOperationException($"file {context.Path} is encrypted");
			}
		}
    }
}
