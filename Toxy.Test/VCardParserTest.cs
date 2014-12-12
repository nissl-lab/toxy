using Microsoft.VisualStudio.TestTools.UnitTesting;

using Toxy.Parsers;

namespace Toxy.Test
{
    [TestClass]
    public class VCardParserTest
    {
        [TestMethod]
        public void TestRead2Cards()
        {
            string path = TestDataSample.GetVCardPath("RfcAuthors.vcf");
            ParserContext context = new ParserContext(path);
            VCardDocumentParser parser = new VCardDocumentParser(context);
            var cards= parser.Parse();
            Assert.AreEqual(2, cards.Cards.Count);

            ToxyBusinessCard tbc1 = cards.Cards[0];
            Assert.AreEqual("Frank Dawson", tbc1.Name.FullName);
            Assert.AreEqual(1, tbc1.Addresses.Count);
            Assert.AreEqual(5, tbc1.Contacts.Count);

            Assert.AreEqual("6544 Battleford Drive;Raleigh;NC;27613-3502;U.S.A.", tbc1.Addresses[0].ToString());
            int i = 0;
            Assert.AreEqual("+1-919-676-9515", tbc1.Contacts[i].Value);
            Assert.AreEqual("MessagingService, WorkVoice", tbc1.Contacts[i].Name);
            i++;
            Assert.AreEqual("+1-919-676-9564", tbc1.Contacts[i].Value);
            Assert.AreEqual("WorkFax", tbc1.Contacts[i].Name);
            i++;
            Assert.AreEqual("Frank_Dawson@Lotus.com", tbc1.Contacts[i].Value);
            Assert.AreEqual("Internet", tbc1.Contacts[i].Name);
            i++;
            Assert.AreEqual("fdawson@earthlink.net", tbc1.Contacts[i].Value);
            Assert.AreEqual("Internet", tbc1.Contacts[i].Name);
            i++;
            Assert.AreEqual("http://home.earthlink.net/~fdawson", tbc1.Contacts[i].Value);
            Assert.AreEqual("Url-Default", tbc1.Contacts[i].Name);
            Assert.AreEqual("Lotus Development Corporation", tbc1.Orgnization);

            //2ed guy
            ToxyBusinessCard tbc2 = cards.Cards[1];
            Assert.AreEqual("Tim Howes", tbc2.Name.FullName);
            Assert.AreEqual("Netscape Communications Corp.", tbc2.Orgnization);
            Assert.AreEqual(1, tbc2.Addresses.Count);
            Assert.AreEqual(3, tbc2.Contacts.Count);
            Assert.AreEqual("501 E. Middlefield Rd.;Mountain View;CA;94043;U.S.A.", tbc2.Addresses[0].ToString());
            i = 0;
            Assert.AreEqual("+1-415-937-3419", tbc2.Contacts[i].Value);
            Assert.AreEqual("MessagingService, WorkVoice", tbc2.Contacts[i].Name);
            i++;
            Assert.AreEqual("+1-415-528-4164", tbc2.Contacts[i].Value);
            Assert.AreEqual("WorkFax", tbc2.Contacts[i].Name);
            i++;
            Assert.AreEqual("howes@netscape.com", tbc2.Contacts[i].Value);
            Assert.AreEqual("Internet", tbc2.Contacts[i].Name);

        }

        [TestMethod]
        public void TestUTF8Card()
        {
            string path = TestDataSample.GetVCardPath("UnicodeNameSample.vcf");
            ParserContext context = new ParserContext(path);
            VCardDocumentParser parser = new VCardDocumentParser(context);
            var cards = parser.Parse();
            Assert.AreEqual(1, cards.Cards.Count);

            var tbc1 = cards.Cards[0];
            Assert.AreEqual(0, tbc1.Addresses.Count);
            Assert.AreEqual(1, tbc1.Contacts.Count);

            Assert.AreEqual("陈丽君", tbc1.Name.FirstName);
            Assert.AreEqual("18777777719", tbc1.Contacts[0].Value);
            Assert.AreEqual("Preferred, CellularVoice", tbc1.Contacts[0].Name);
        }

        [TestMethod]
        public void TestForeignNames()
        {
            string path = TestDataSample.GetVCardPath("PalmAgentSamples.vcf");
            ParserContext context = new ParserContext(path);
            VCardDocumentParser parser = new VCardDocumentParser(context);
            var source = parser.Parse();
            Assert.AreEqual(20,source.Cards.Count);

            Assert.AreEqual("John Doe4", source.Cards[10].Name.FullName);
        }
    }
}
