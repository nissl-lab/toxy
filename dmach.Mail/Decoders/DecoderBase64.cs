using System;
using System.IO;
using System.Collections;

namespace InSolve.dmach.Mail.Decoders
{
	/// <summary>
	/// Base64 Decoder
	/// </summary>
	public class DecoderBase64 
		: DecoderBase
	{
		#region Fields

		private Queue						_queue;
		private static readonly byte[]		_symbols = GetSymbols();

		#endregion
		
		#region Constructors

		public DecoderBase64()
			:base()
		{
			_queue	= new Queue(8);
		}

		#endregion

		#region DecoderBase Members

		public override int Add(byte[] b, int start, int length)
		{
			lock(StreamSyncRoot)
			{
				int r = 0;	//return value
				for(int i=0; i<length; i++)
				{
					byte sym = b[start+i];

					if (ContainsSymbol(sym))
					{
						_queue.Enqueue(_symbols[sym]);
						if (_queue.Count == 4)
						{
							unchecked
							{
								//алгоритм получения 3-х реальных байт из 4-х байт Base64
								byte[] t = new byte[4];
								for(int k=0; k<4; k++)
									t[k] = (byte)_queue.Dequeue();
						
								InnerBuffer.WriteByte((byte)((t[0]<<2)|(t[1]>>4)));
								r++;
								if (!IsEqualSign(t[2]))
								{
									InnerBuffer.WriteByte((byte)((t[1]<<4)|(t[2]>>2)));
									r++;
								}
								if (!IsEqualSign(t[3]))
								{
									InnerBuffer.WriteByte((byte)((t[2]<<6)|t[3]));
									r++;
								}
							}
						}
					}
				}
				return r;
			}
		}

		#endregion

		#region Tool Methods

		private static byte[] GetSymbols()
		{
			byte[] b = new byte[256];
			for(int i = 0; i < b.Length; i++)
				b[i] = 255;	//error fields

			string pattern = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
			for(byte i = 0; i < pattern.Length; i++)
				b[(byte)pattern[i]] = i;
			b[(byte)'='] = 254;	//'особый разговор' - было написано в прежнем алгоритме.

			return b;
		}

		private static bool ContainsSymbol(byte sym)
		{
			return (_symbols[sym] == 255) ? false : true;
		}

		private static bool IsEqualSign(byte sym)
		{
			return (sym == 254);
		}

		#endregion
	}
}
