using System;
using System.IO;

namespace InSolve.dmach.Mail.Decoders
{
	public class Decoder8bit : 
		DecoderBase
	{
		#region Constructors

		public Decoder8bit()
			:base()
		{}

		#endregion

		#region DecoderBase Members

		public override int Add(byte[] b, int start, int length)
		{
			lock(StreamSyncRoot)
			{
				InnerBuffer.Write(b, start, length);
				return length;
			}
		}

		#endregion
	}
}
