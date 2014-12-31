using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class ImageParserTest
    {
        [Test]
        public void TestParseJpeg()
        {
            string path = Path.GetFullPath(TestDataSample.GetImagePath("sample_sony1.jpg"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(12, x.Count);
            Assert.AreEqual(2592, x.Get("PhotoHeight").Value);
            Assert.AreEqual(95, x.Get("PhotoQuality").Value);
            Assert.AreEqual(3872, x.Get("PhotoWidth").Value);
            Assert.AreEqual("DSLR-A200", x.Get("Model").Value);
            Assert.AreEqual((uint)400, x.Get("ISOSpeedRatings").Value );
            Assert.AreEqual(5.6, x.Get("FNumber").Value);
            Assert.AreEqual((double)35, x.Get("FocalLength").Value );
            Assert.AreEqual((uint)52, x.Get("FocalLengthIn35mmFilm").Value );
            Assert.AreEqual(new DateTime(2009, 11, 21, 12, 39, 39), x.Get("DateTime").Value);
        }
        [Test]
        public void TestParseJpegWithXmp()
        {
            string path = Path.GetFullPath(TestDataSample.GetImagePath("sample_nikon1.jpg"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(16, x.Count);
            Assert.AreEqual(3008, x.Get("PhotoHeight").Value);
            Assert.AreEqual(96, x.Get("PhotoQuality").Value);
            Assert.AreEqual(2000, x.Get("PhotoWidth").Value);
            Assert.AreEqual("NIKON D70", x.Get("Model").Value);
            Assert.AreEqual("Kirche Sulzbach", x.Get("Keywords").Value);
            Assert.AreEqual("1", x.Get("Rating").Value);
            Assert.AreEqual("2009-08-04T20:42:36Z", x.Get("DateAcquired").Value);
            Assert.AreEqual("Microsoft Windows Photo Gallery 6.0.6001.18000", x.Get("Software").Value);
            Assert.AreEqual("Microsoft Windows Photo Gallery 6.0.6001.18000", x.Get("creatortool").Value);
        }
        [Test]
        public void TestParseGifWithComment()
        {
            string path = Path.GetFullPath(TestDataSample.GetImagePath("sample_gimp.gif"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(4, x.Count);
            Assert.AreEqual(37, x.Get("PhotoHeight").Value);
            Assert.AreEqual(12, x.Get("PhotoWidth").Value);
            Assert.AreEqual("Created with GIMP", x.Get("Comment").Value);
            Assert.AreEqual("Created with GIMP", x.Get("GifComment").Value);
        }
        [Test]
        public void TestParseTiff()
        {
            string path = Path.GetFullPath(TestDataSample.GetImagePath("sample_gimp.tiff"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(97, x.Count);
            Assert.AreEqual(10, x.Get("PhotoHeight").Value);
            Assert.AreEqual(10, x.Get("PhotoWidth").Value);
            Assert.AreEqual("Test", x.Get("Comment").Value);
            Assert.AreEqual("28/10", x.Get("FNumber").Value);
            Assert.AreEqual("0", x.Get("Rating").Value);
            Assert.AreEqual("+150", x.Get("Tint").Value);
            Assert.AreEqual("5", x.Get("Shadows").Value);
        }
        
    }
}
