using System;
using System.Text;

namespace InSolve.dmach.Mail
{
	internal class StreamStringBuilder
	{
		private StringBuilder _buffer;

		public StreamStringBuilder()
			:this(128)
		{}

		public StreamStringBuilder(int sizeOfInnerBuffer)
		{
			_buffer = new StringBuilder(sizeOfInnerBuffer);
		}

		public void Add(byte symbol)
		{
			_buffer.Append(
				(char)((symbol > 191)? symbol + 848 : symbol)
				);
		}

		public int Add(byte[] buffer, int offset, int count)
		{
			int i = 0;
			while(count-- > 0)
			{
				Add(buffer[offset++]);
				i++;
			}

			return i;
		}

		public override string ToString()
		{
			return _buffer.ToString();
		}

		public void Reset()
		{
			_buffer.Length = 0;
		}
	}
}
