using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InSolve.dmach.Mail
{
    public enum AttachmentFormat
    {
        Unknown,
        PlainText,
        Html,
        ByteStream,
        /// <summary>
        /// Вложенное письмо
        /// </summary>
        ByteStreamMessage,
    }
}
