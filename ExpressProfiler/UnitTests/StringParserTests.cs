using System;
using ExpressProfiler;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    public class StringParserTests
    {
        [Test]
        public void TestSplitOnSpace()
        {
            const string s =
                "exec sp_executesql N'SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@Table',N'@Table nvarchar(11)',@Table=N'AspNetUsers'";

            var items = StringParser.Split(s, ' ', '\'');

            Assert.AreEqual(3, items.Count);
            Assert.AreEqual("exec", items[0]);
            Assert.AreEqual("sp_executesql", items[1]);
            StringAssert.StartsWith("N'SELECT COLUMN_NAME ", items[2]);
        }

        [Test]
        public void TestSplitOnComma()
        {
            const string s =
                "N'SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@Table',N'@Table nvarchar(11)',@Table=N'AspNetUsers'";

            var items = StringParser.Split(s, ',', '\'');

            Assert.AreEqual(3, items.Count);
            Assert.AreEqual("N'SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@Table'", items[0]);
            Assert.AreEqual("N'@Table nvarchar(11)'", items[1]);
            Assert.AreEqual("@Table=N'AspNetUsers'", items[2]);
        }

    }
}