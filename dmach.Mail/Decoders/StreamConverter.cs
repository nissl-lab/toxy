using System;
using System.IO;

namespace InSolve.dmach.Mail.Decoders
{
	/// <summary>
	/// На лету кодирует данные блоками по 4096 байт декодером к указанный поток.
	/// </summary>
	internal class StreamConverter :
		Stream
	{
		#region Fields

		private Stream		_targetStream;
		private DecoderBase	_decoder;

		public StreamConverter(Stream targetStream, DecoderBase decoder)
		{
			if (targetStream == null)
				throw new ArgumentNullException("targetStream");

			if (decoder == null)
				throw new ArgumentNullException("decoder");

			_targetStream	= targetStream;
			_decoder		= decoder;
		}

		#endregion

		#region Override abstract System.IO.Stream members

		public override bool CanRead
		{
			get
			{
				return false;
			}
		}

		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		public override bool CanWrite
		{
			get
			{
				return true;
			}
		}

		public override void Flush()
		{
			if (_targetStream == null)
				throw new InvalidOperationException("Stream closed");

			_decoder.WriteTo(_targetStream);
			_targetStream.Flush();
		}

		public override long Length
		{
			get
			{
				throw new NotSupportedException("Length get");
			}
		}

		public override long Position
		{
			get
			{
				throw new NotSupportedException("Position get");
			}
			set
			{
				throw new NotSupportedException("Position set");
			}
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("Read()");
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("Seek()");
		}

		public override void SetLength(long value)
		{
			throw new NotSupportedException("SetLength()");
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
			if (_targetStream == null)
				throw new InvalidOperationException("Stream closed");

			int bufferSize = 4096;

			while(count > 0)
			{
				_decoder.Add(
					buffer,
					offset, 
					(count <= bufferSize)? count : bufferSize
					);

				count -= bufferSize;
				offset += bufferSize;
				
				_decoder.WriteTo(_targetStream);
			}

			return;
		}

		#endregion

		#region Override virtual base class members

		public override void Close()
		{
			Flush();
			_targetStream	= null;
			_decoder		= null;

			base.Close();
		}

		#endregion

		#region Methods

		public Stream GetTargetStream()
		{
			Flush();
			return _targetStream;
		}

		#endregion

		#region Static methods

		public static int Convert(Stream inStream, int innerBufferSize, DecoderBase decoder, Stream outStream)
		{
			#region Check parameters

			if (inStream == null)
				throw new ArgumentNullException("inStream");
			if (innerBufferSize < 1)
				throw new ArgumentException("innerBufferSize less zero");
			if (decoder == null)
				throw new ArgumentNullException("decoder");
			if (outStream == null)
				throw new ArgumentNullException("outStream");

			#endregion

			ParserStreamPos pos = new ParserStreamPos(innerBufferSize);
			int count = 0;

			for(;;)
			{
				if (pos.IsFillNeed)
				{
					pos.Read(inStream);
					if (pos.EndOfStream)
						break;
				}

				count += decoder.Add(
					pos.Buffer,
					pos.CurrentPosition,
					pos.GetLengthToEnd()
					);

				pos.CurrentPosition += pos.GetLengthToEnd();
				decoder.WriteTo(outStream);
			}

			return count;
		}

		#endregion
	}
}
