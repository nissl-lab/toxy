using System;
using System.IO;

namespace InSolve.dmach.Mail
{
	/// <summary>
	/// —труктура описывающа€ состо€ние проверки
	/// </summary>
	internal struct ParserStreamPos
	{
		#region Fields

		public byte[]	Buffer;

		public int		CurrentPosition;
		public int		BufferSize;

		private bool	_endOfStream;

		#endregion

		public ParserStreamPos(int sizeOfInnerBuffer)
		{
			CurrentPosition	= 0;
			BufferSize		= 0;
			_endOfStream	= false;

			Buffer = new byte[sizeOfInnerBuffer];
		}

		public void Close()
		{
			Buffer = null;
		}

		public bool IsFillNeed
		{
			get
			{
				return GetLengthToEnd() <= 0;
			}
		}

		public bool EndOfStream
		{
			get
			{
				return _endOfStream;
			}
		}

		public int Read(Stream stream)
		{
			if (stream == null)
				throw new ArgumentNullException("stream");

			CurrentPosition = 0;
			BufferSize = stream.Read(Buffer, 0, Buffer.Length);
			_endOfStream = (BufferSize == 0);

			return BufferSize;
		}

		public int GetIndexOf(int symbol)
		{
			if (symbol == -1)
				return -1;

			if (!IsFillNeed)
				return Array.IndexOf(Buffer, (byte)symbol, CurrentPosition, GetLengthToEnd());

			throw new InvalidOperationException("inner buffer is empty");
		}

		public int GetLengthToEnd()
		{
			return BufferSize - CurrentPosition;
		}

		/// <summary>
		/// Length to index (except), if index is -1 method return it
		/// </summary>
		/// <param name="indexOfSymbol"></param>
		/// <returns></returns>
		public int GetLengthToIndex(int indexOfSymbol)
		{
			return (indexOfSymbol != -1)? indexOfSymbol - CurrentPosition : -1;
		}

	}
}
