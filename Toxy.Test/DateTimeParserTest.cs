using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;

namespace Toxy.Test
{
    [TestFixture]
    public class DateTimeParserTest
    {
        [Ignore("to be fixed")]
        public void TestParseDatetimeWithTimezone()
        {
            ClassicAssert.AreEqual(new DateTime(2008, 10, 25, 3, 9, 6, DateTimeKind.Utc), DateTimeParser.Parse("24-oct-08 21:09:06 CEST"));
            ClassicAssert.AreEqual(new DateTime(2012, 4, 20, 16, 10, 0, DateTimeKind.Utc), DateTimeParser.Parse("2012-04-20 10:10:00+0200"));
            ClassicAssert.AreEqual(new DateTime(2014, 12, 12, 12, 13, 30, DateTimeKind.Utc), DateTimeParser.Parse("Fri, 12 Dec 2014 12:13:30 +0800 (CST)"));
        }

        [Test]
        public void TestParseDatetimeWithoutTimezone()
        {
            ClassicAssert.AreEqual(new DateTime(2014, 12, 12, 12, 13, 30), DateTimeParser.Parse("Fri, 12 Dec 2014 12:13:30"));
        }
    }
}
