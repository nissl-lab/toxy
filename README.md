[![NuGet Badge](https://buildstats.info/nuget/Toxy)](https://www.nuget.org/packages/Toxy)
[![Ko-Fi](https://img.shields.io/static/v1?style=flat-square&message=Support%20the%20Project&color=success&style=plastic&logo=ko-fi&label=$$)](https://ko-fi.com/tonyqus)
[![netstandard2.0](https://img.shields.io/badge/netstandard-2.0-brightgreen.svg)](https://img.shields.io/badge/netstandard-2.0-brightgreen.svg)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg?style=flat-square&logo=Apache)](License.md)

What's Toxy
============
Toxy is a .NET data/text extraction framework similar to Apache Tika in Java. It supports a lot of popular formats such as docx, xlsx, xls, pdf, csv, txt, epub, html and so on.

![image](https://user-images.githubusercontent.com/772561/131231873-e22b4170-1dd5-4e35-b928-7732c80065ea.png)


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
