using System;
using System.IO;

namespace InSolve.dmach.Mail
{
	internal abstract class StreamTools
	{
		public static byte[] GetBytesFromBegin(Stream stream)
		{
			if (stream == null)
				return new byte[0];

			if (stream is MemoryStream)
				return (stream as MemoryStream).ToArray();

			if (!stream.CanSeek)
				throw new NotSupportedException("stream must be 'CanSeek'");

			if (stream.Position != 0)
				stream.Seek(0, SeekOrigin.Begin);

			MemoryStream buffer = new MemoryStream((int)stream.Length);

			for(byte[] bytes = new byte[4096];;)
			{
				int count = stream.Read(bytes, 0, 4096);

				if (count == 0)
					break;

				buffer.Write(bytes, 0, count);
			}
			
			return buffer.ToArray();
		}	
	}
}
