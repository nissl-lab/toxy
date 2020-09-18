//Сделано 14.06.2004

using System;
using System.IO;
using System.Text;

namespace InSolve.dmach.Mail
{
	/// <summary>
	/// Класс для парсинга почты.
	/// </summary>
	public abstract class Parser
	{
		#region Static fields and buffers
		
		/// <summary>
		/// Структура описывающая положение в потоке и т.п.
		/// </summary>
		[ThreadStatic]
		private static ParserStreamPos	_pos;
		
		#endregion

		#region Stream read methods

		/// <summary>
		/// Возвращает строку из потока. Строка от начала до первого \r\n _включительно_
		/// </summary>
		/// <param name="stream">поток</param>
		/// <param name="buff">буфер</param>
		/// <returns></returns>
		private static string ReadStreamLine(Stream stream)
		{
			//Алгоритм проверен, работает классно!
			//(на минимальных, максимальных и средних значениях длины буфера)
			if (_pos.EndOfStream)
				return string.Empty;

			StreamStringBuilder buffer = new StreamStringBuilder();

			for(;;)
			{
				if (_pos.IsFillNeed)
				{
					_pos.Read(stream);
					if (_pos.EndOfStream)
						break;
				}

				int indexOfEnter = _pos.GetIndexOf((int)'\n');

				_pos.CurrentPosition += buffer.Add(
					_pos.Buffer,
					_pos.CurrentPosition,
					(indexOfEnter == -1)? _pos.GetLengthToEnd() : _pos.GetLengthToIndex(indexOfEnter) + 1
					);

				if (indexOfEnter != -1)
					break;
			}

			return buffer.ToString();
		}

		/// <summary>
		/// Читает данные до указанного разделителя не включая его в результат.
		/// </summary>
		/// <param name="outStream"></param>
		/// <param name="stream"></param>
		/// <param name="letterError"></param>
		/// <param name="currentBound"></param>
		/// <param name="finalPart"></param>
		private static void ReadStreamToBound
			(
			Stream outStream,
			Stream stream,
            Action<ParserError, string> errorHandler,
			string currentBound,
			out bool finalPart
			)
		{
			//Алгоритм проверен, работает классно!
			//(на минимальных, максимальных и средних значениях длины буфера)
			//на большом, малом и вообще без разделителя
			//на вариантах для разделителя "bound": bbbound
			//известный глюк: строку вида --a в подстроке вида ---a она не должна найти. Это фича :-)
			byte[] bound = ParserTools.GetBytes(currentBound);

			for(int progress = 0;;)
			{
				if (bound != null && bound.Length == progress)
					break;

				if (_pos.IsFillNeed)
				{
					_pos.Read(stream);
					if (_pos.EndOfStream)
					{
						if (progress > 0)
							if (outStream != null)
								outStream.Write(bound, 0, progress);
						break;
					}
				}

				int count = _pos.GetLengthToIndex(
					(bound != null)? _pos.GetIndexOf(bound[progress]) : -1
					);

				if (count == 0)
				{
					progress++;
					_pos.CurrentPosition++;
					continue;
				}
				if (progress > 0)
				{
					if (outStream != null)
						outStream.Write(bound, 0, progress);
					progress = 0;
					continue;
				}
				if (count < 0)
					count = _pos.GetLengthToEnd();
				
				if (outStream != null)
					outStream.Write(_pos.Buffer, _pos.CurrentPosition, count);
				_pos.CurrentPosition += count;
			}

			if (_pos.EndOfStream)
			{
                if (currentBound != null && errorHandler != null)
                    errorHandler(ParserError.NoCorrectBoundInMessageEnd, null);
                
                finalPart = true;                
                return;
			}

			string str = Parser.ReadStreamLine(stream);
			finalPart = (str.IndexOf("--") != -1);

            return;
		}

		#endregion

        /// <summary>
        /// Внутреннее представление парсера, конструирует из потока IMessageConstructor
        /// </summary>
        /// <param name="constructor">Экземпляр IMessageConstructor</param>
        /// <param name="stream">Поток для чтения данных</param>
        /// <param name="sizeOfInnerBuffer">Размер буфера, например 4096</param>
        /// <param name="upperLevel">Верхний уровень рекурсии, при первом запуске обязан быть true</param>
        /// <param name="currentBound">Параметр для внутренней рекурсии, обязан быть null при первом запуске</param>
        /// <returns>finalPart</returns>
		static bool Parse
			(
			IMessageConstructor constructor,
			Stream stream,
			int sizeOfInnerBuffer,
			bool upperLevel,
			string currentBound
			)
		{
			#region Check parameters

			if (constructor == null)
				throw new ArgumentNullException("constructor");
			if (stream == null)
				throw new ArgumentNullException("stream");

			if (upperLevel)
			{
				//создать новый экземпляр жизненно необходимо для того, что бы затереть данные предыдущего парсера в этом потоке
				_pos = new ParserStreamPos(sizeOfInnerBuffer);
				_pos.Read(stream);
			}

            Action<ParserError, string> errorHandler = (err, comment) => constructor.ErrorHandler(err, comment);

			#endregion

			#region Create variable

			HeaderCollection	headers		= new HeaderCollection();
			bool				finalPart	;
			Encoding			encoding	= Encoding.Default;
			string				nextBound	= null;

			#endregion

			#region Headers parse
			
			StringBuilder strBuffer = new StringBuilder(76*3);	//76 - длина строки в письме по RFC822
			
			for(;;)
			{
				string str = Parser.ReadStreamLine(stream).Trim();

				if (_pos.EndOfStream)
				{
                    errorHandler(ParserError.NoContainsBodyContent, null);
					
					finalPart = true;
					return finalPart;
				}

				if (str.Length == 0)
				{
					//надо переходить в вариант обработки текста
					if (strBuffer.Length > 0)
                        ParserTools.ParseHeader(strBuffer.ToString(), headers, constructor, errorHandler, upperLevel);
					break;
				}

				//переводим строку в вид без кодировок
				bool correctHeader;
                str = ParserTools.ConvertMimeHeader(str, out correctHeader, errorHandler);
				
				if (correctHeader)
				{
					if (strBuffer.Length > 0)
                        ParserTools.ParseHeader(strBuffer.ToString(), headers, constructor, errorHandler, upperLevel);

					strBuffer.Length = 0;
				}

				strBuffer.Append(str);
			}

			#endregion

			#region Before content parse

			nextBound = ParserTools.GetNextBound(headers);
			//если проверяемая часть многочастная
			if (nextBound != null)
			{
				//игнорируем данные между концом заголовков и началом части
				bool ignoreIt;
                Parser.ReadStreamToBound(null, stream, errorHandler, nextBound, out ignoreIt);
			}

			#endregion

			#region Content parse

			if (nextBound != null)
			{
                while (!Parser.Parse(constructor, stream, sizeOfInnerBuffer, false, nextBound)) ;

				if (_pos.EndOfStream)
					finalPart = true;
				else
                    Parser.ReadStreamToBound(null, stream, errorHandler, currentBound, out finalPart);
			}
			else
			{
				//дополняем разделитель энтером впереди по стандарту
				if (currentBound != null)
					currentBound = Environment.NewLine + currentBound;

				IAttachmentConstructor attachment = constructor.GetAttachmentConstructor(headers);
				
				Parser.ReadStreamToBound(
					(attachment != null)? attachment.Stream : null,
					stream,
                    errorHandler,
					currentBound,
					out finalPart
					);

                if (attachment != null)
                {
                    attachment.CompleteConstruction();
                    constructor.CompleteAttachment(attachment);
                }
			}

			#endregion

			#region After content parse

			if (upperLevel)
			{
				constructor.CompleteConstruction();
				_pos.Close();
			}

			#endregion

			return finalPart;
		}

        public static void Parse(IMessageConstructor constructor, Stream stream)
        {
            Parse(constructor, stream, 4096);
        }

        public static void Parse(IMessageConstructor constructor, Stream stream, int sizeOfInnerBuffer)
        {
            Parse(constructor, stream, sizeOfInnerBuffer, true, null);
        }

	}
}
