using NUnit.Framework;
using NUnit.Framework.Legacy;
using Toxy.Parsers;

namespace Toxy.Test
{
    [TestFixture]
    public class VCardParserTest
    {
        [Test]
        public void TestRead2Cards()
        {
            ParserContext context = new ParserContext(TestDataSample.GetVCardPath("RfcAuthors.vcf"));
            VCardDocumentParser parser = ParserFactory.CreateVCard(context);
            var cards= parser.Parse();
            ClassicAssert.AreEqual(2, cards.Cards.Count);

            ToxyBusinessCard tbc1 = cards.Cards[0];
            ClassicAssert.AreEqual("Frank Dawson", tbc1.Name.FullName);
            ClassicAssert.AreEqual(1, tbc1.Addresses.Count);
            ClassicAssert.AreEqual(5, tbc1.Contacts.Count);

            ClassicAssert.AreEqual("6544 Battleford Drive;Raleigh;NC;27613-3502;U.S.A.", tbc1.Addresses[0].ToString());
            int i = 0;
            ClassicAssert.AreEqual("+1-919-676-9515", tbc1.Contacts[i].Value);
            ClassicAssert.AreEqual("MessagingService, WorkVoice", tbc1.Contacts[i].Name);
            i++;
            ClassicAssert.AreEqual("+1-919-676-9564", tbc1.Contacts[i].Value);
            ClassicAssert.AreEqual("WorkFax", tbc1.Contacts[i].Name);
            i++;
            ClassicAssert.AreEqual("Frank_Dawson@Lotus.com", tbc1.Contacts[i].Value);
            ClassicAssert.AreEqual("Internet", tbc1.Contacts[i].Name);
            i++;
            ClassicAssert.AreEqual("fdawson@earthlink.net", tbc1.Contacts[i].Value);
            ClassicAssert.AreEqual("Internet", tbc1.Contacts[i].Name);
            i++;
            ClassicAssert.AreEqual("http://home.earthlink.net/~fdawson", tbc1.Contacts[i].Value);
            ClassicAssert.AreEqual("Url-Default", tbc1.Contacts[i].Name);
            ClassicAssert.AreEqual("Lotus Development Corporation", tbc1.Orgnization);

            //2ed guy
            ToxyBusinessCard tbc2 = cards.Cards[1];
            ClassicAssert.AreEqual("Tim Howes", tbc2.Name.FullName);
            ClassicAssert.AreEqual("Netscape Communications Corp.", tbc2.Orgnization);
            ClassicAssert.AreEqual(1, tbc2.Addresses.Count);
            ClassicAssert.AreEqual(3, tbc2.Contacts.Count);
            ClassicAssert.AreEqual("501 E. Middlefield Rd.;Mountain View;CA;94043;U.S.A.", tbc2.Addresses[0].ToString());
            i = 0;
            ClassicAssert.AreEqual("+1-415-937-3419", tbc2.Contacts[i].Value);
            ClassicAssert.AreEqual("MessagingService, WorkVoice", tbc2.Contacts[i].Name);
            i++;
            ClassicAssert.AreEqual("+1-415-528-4164", tbc2.Contacts[i].Value);
            ClassicAssert.AreEqual("WorkFax", tbc2.Contacts[i].Name);
            i++;
            ClassicAssert.AreEqual("howes@netscape.com", tbc2.Contacts[i].Value);
            ClassicAssert.AreEqual("Internet", tbc2.Contacts[i].Name);

        }

        [Test]
        public void TestUTF8Card()
        {
            ParserContext context = new ParserContext(TestDataSample.GetVCardPath("UnicodeNameSample.vcf"));
            VCardDocumentParser parser = ParserFactory.CreateVCard(context);
            var cards = parser.Parse();
            ClassicAssert.AreEqual(1, cards.Cards.Count);

            var tbc1 = cards.Cards[0];
            ClassicAssert.AreEqual(0, tbc1.Addresses.Count);
            ClassicAssert.AreEqual(1, tbc1.Contacts.Count);

            ClassicAssert.AreEqual("陈丽君", tbc1.Name.FirstName);
            ClassicAssert.AreEqual("18777777719", tbc1.Contacts[0].Value);
            ClassicAssert.AreEqual("Preferred, CellularVoice", tbc1.Contacts[0].Name);
        }

        [Test]
        public void TestForeignNames()
        {
            string path = TestDataSample.GetVCardPath("PalmAgentSamples.vcf");
            ParserContext context = new ParserContext(path);
            VCardDocumentParser parser = ParserFactory.CreateVCard(context);
            var source = parser.Parse();
            ClassicAssert.AreEqual(20,source.Cards.Count);

            ClassicAssert.AreEqual("John Doe4", source.Cards[10].Name.FullName);
        }
    }
}
