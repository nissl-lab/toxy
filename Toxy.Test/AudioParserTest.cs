using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class AudioParserTest
    {
        [Test]
        public void TestParseMp3()
        {
            string path = Path.GetFullPath(TestDataSample.GetAudioPath("sample_both.mp3"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(15, x.Count);
        }
        [Test]
        public void TestParseMp3_Id3v1Only()
        {
            string path = Path.GetFullPath(TestDataSample.GetAudioPath("sample_v1_only.mp3"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(11, x.Count);
        }
        [Test]
        public void TestParseMp3_Id3v2Only()
        {
            string path = Path.GetFullPath(TestDataSample.GetAudioPath("sample_v2_only.mp3"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(15, x.Count);
        }
        [Test]
        public void TestParseApe()
        {
            string path = Path.GetFullPath(TestDataSample.GetAudioPath("sample.ape"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(15, x.Count);

        }
        [Test]
        public void TestParseWma()
        {
            string path = Path.GetFullPath(TestDataSample.GetAudioPath("sample.wma"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(17, x.Count);
        }
        [Test]
        public void TestParseFlac()
        {
            string path = Path.GetFullPath(TestDataSample.GetAudioPath("sample.flac"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(16, x.Count);
        }
    }
}
