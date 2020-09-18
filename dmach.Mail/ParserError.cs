using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InSolve.dmach.Mail
{
    [Flags]
    public enum ParserError
    {
        /// <summary>
        /// Письмо не завершилось корректным разделителем
        /// </summary>
        NoCorrectBoundInMessageEnd = 1,
        /// <summary>
        /// Не удалось получить нужный декодер (группа transfer) заголовка из строки
        /// </summary>
        FailGetDecoderFromMimeString = 2,
        /// <summary>
        /// Заголовок не соответствует mime-шаблону
        /// </summary>
        HeaderMismatchMimePattern = 4,
        /// <summary>
        /// Письмо закончилось на заголовках
        /// </summary>
        NoContainsBodyContent = 8,
        /// <summary>
        /// Неизвестная кодировка, используется кодировка по умолчанию
        /// </summary>
        UnknownEncoding = 16,
        /// <summary>
        /// Не получилось создать корректный декодер для текста (7bit, 8bit, base64, q-p, etc...)
        /// </summary>
        FailGetDecoderByPattern = 32,


    }
}
