using System;
using System.Collections.Generic;

namespace InSolve.dmach.Mail.Decoders
{
	/// <summary>
	/// Позволяет получить нужный декодер. Подразумевается, что в будущем человеки смогут наследоваться
	/// от, и вводить в игру свои декодеры.
	/// </summary>
	public class DecoderSelector
	{
        class RegisteredDecoder
        {
            public string Pattern {get; set;}
            public Func<DecoderBase> Decoder {get; set;}
        }

        List<RegisteredDecoder> RegDecoders = new List<RegisteredDecoder>();

		#region Constructors

		public DecoderSelector()
		{
			// Register well-know decoders

            RegisterDecoder("quoted", () => new DecoderQP());
            RegisterDecoder("64", () => new DecoderBase64());
            RegisterDecoder("8", () => new Decoder8bit());
            RegisterDecoder("7", () => new Decoder7bit());
        }

		#endregion

		#region Members

        public void RegisterDecoder(string pattern, Func<DecoderBase> decoder)
        {
            if (decoder == null)
                throw new ArgumentNullException("decoder");

            RegDecoders.Insert(0, new RegisteredDecoder() { Pattern = pattern, Decoder = decoder });
        }

		public DecoderBase GetDecoder(string desc, Action<ParserError, string> errorHandler)
		{
            DecoderBase d = null;
            if (desc != null)
            {
                foreach (var r in RegDecoders)
                {
                    if (desc.IndexOf(r.Pattern, StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        d = r.Decoder();
                        break;
                    }
                }
            }

            if (d == null)
            {
                if (desc != null && errorHandler != null)
                    errorHandler(ParserError.FailGetDecoderByPattern, desc);

                d = new Decoder7bit();
            }

			return d;
		}

        public DecoderBase GetDecoder(HeaderCollection headers, Action<ParserError, string> errorHandler)
		{
			if (headers == null)
				throw new ArgumentNullException("headers");

			return GetDecoder(headers.Get(Mime.Content_Transfer_Encoding), errorHandler);
		}

		#endregion

		/// <summary>
		/// Назначает декодер для раскодирования заголовка
		/// </summary>
		/// <param name="desc">буква декодера</param>
		/// <param name="decSelector">селектор декодеров</param>
		/// <param name="decoder">out: выбранный декодер</param>
		/// <returns>уверенно назначен декодер или нет</returns>
		internal static DecoderBase GetHeaderDecoder(string desc, out bool confidently)
		{
			if (desc == null)
				throw new ArgumentNullException("desc");
			
			desc = desc.ToLower().Trim();

			switch(desc)
			{
				case "b":
					confidently = true;
					return new DecoderBase64();
				case "q":
					confidently = true;
					return new DecoderQP();
				default:
					confidently = false;
					return new Decoder7bit();
			}
		}

	}
}
