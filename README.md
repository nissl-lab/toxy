Who are We
==========

[![Join the chat at https://gitter.im/tonyqus/toxy](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/tonyqus/toxy?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

What's Toxy
============
Toxy is a .NET data/text extraction framework similar to Apache Tika in Java. It supports a lot of popular formats such as docx, xlsx, xls, pdf, csv, txt, epub, html and so on.

Why Toxy
============
In the past, we have to use IFilter to extract texts from other documents. But Toxy is platform independent. It will try to support not only Windows but also Linux (with Mono installed). The usage of Toxy will be very easy. You don't need to care much about what extension you are extracting because it will create a clever framework to help identify the file formats and extract the data/text into a unified structure. 

Toxy Objects
==================
- ToxyDocument - the data structure extracted for documents
- ToxySpreadsheet - the data structure extracted for spreadsheets
- ToxyEmail - the data structure extracted for emails
- ToxyBusinessCard - the data structure extracted for business cards
- ToxyDom - the data structure extracted for DOM based document
- ToxyMetadata - the data structure extracted for other files with meta data
