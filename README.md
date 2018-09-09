Who are We
==========

[![Join the chat at https://gitter.im/tonyqus/toxy](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/tonyqus/toxy?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)
Neuzilla is the studio behind Toxy. For detail, you can check http://www.linkedin.com/company/neuzilla 

What's Toxy
============
Toxy is a .NET data/text extraction framework similar to Apache Tika in Java. It supports a lot of popular formats such as docx, xlsx, xls, pdf, csv, txt, epub, html and so on.

Why Toxy
============
In the past, we have to use IFilter to extract texts from other documents. But Toxy is platform independent. It will try to support not only Windows but also Linux (with Mono installed). The usage of Toxy will be very easy. You don't need to care much about what extension you are extracting because it will create a clever framework to help identify the file formats and extract the data/text into a unified structure. 

For documents, the data structure is called ToxyDocument.
For spreadsheets, the data structure is called ToxySpreadsheet.
For emails, the data structure is called ToxyEmail.
For business cards, the data structure is called ToxyBusinessCard.
For DOM based structure, the data structue is called ToxyDom.
For metadata, the data structure is called ToxyMetadata (Since Toxy 1.3)


.NET Core support
============
I have ported Toxy with most features to .NET Core (PDF, doc, docx, xls, xlsx, vCard, email ..etc, NO JPEG and Video meta extract support), Some project does't have .NET Core verison, I migration it or replaced it with another library

Hear is a list of some .NET Core unsupported library:

NPOI.ScratchPad.HWPF.dll
Thought.vCards.dll
dmach.Mail
iTextSharp (PDF about)
DCSoft.RTF

I have not full tested it, but I will fix the bug in the feature.

Hope it can help you.


How to join
============
If you have passion to contribute code, please contact tony.qu@neuzilla.com

Contact Neuzilla
============
Homepage: http://blog.neuzilla.com
Linkedin User Group: http://www.linkedin.com/groups/Neuzilla-User-Group-6655065 
QQ Group:297128022
Twitter: http://twitter.com/tonyqus
Email: support@neuzilla.com

Latest Source Code
====================
Github: https://github.com/tonyqus/toxy
