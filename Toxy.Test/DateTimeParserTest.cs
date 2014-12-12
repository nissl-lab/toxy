using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestClass]
    public class DateTimeParserTest
    {
        [TestMethod]
        public void TestParseDatetimeWithTimezone()
        {
            Assert.AreEqual(new DateTime(2008, 10, 25, 3, 9, 6), DateTimeParser.Parse("24-oct-08 21:09:06 CEST"));
            Assert.AreEqual(new DateTime(2012, 4, 20, 16, 10, 0), DateTimeParser.Parse("2012-04-20 10:10:00+0200"));
            Assert.AreEqual(new DateTime(2014, 12, 12, 12, 13, 30), DateTimeParser.Parse("Fri, 12 Dec 2014 12:13:30 +0800 (CST)"));

        }

        [TestMethod]
        public void TestParseDatetimeWithoutTimezone()
        {
            Assert.AreEqual(new DateTime(2014, 12, 12, 12, 13, 30), DateTimeParser.Parse("Fri, 12 Dec 2014 12:13:30"));
        }
    }
}
