using System;
using ExpressProfiler;
using NUnit.Framework;
using NUnit.Framework.Internal;

namespace UnitTests
{
    [TestFixture]
    public class SqlFormatterTests
    {
        
        // Poor Man's Formatter thinks the sql in the following is an error. SSMS doesn't though.
        //
        // exec sp_executesql N'BEGIN CONVERSATION TIMER (''de9f14b1-254b-eb11-ae66-983b8fd5df35'') TIMEOUT = 120; WAITFOR(RECEIVE TOP (1) message_type_name, conversation_handle, cast(message_body AS XML) as message_body from [SqlQueryNotificationService-5489eb19-dd3e-4d99-bd14-a59c48e22a2d]), TIMEOUT @p2;',N'@p2 int',@p2=60000

        [TestCase("abc", "abc")]
        [TestCase("abc", "'abc'")]
        [TestCase("abc", "N'abc'")]
        [TestCase("abc", " N'abc' ")]
        public void TestRemovingSqlQuotes(string expected, string input)
        {
            Assert.AreEqual(expected, SqlFormatter.RemoveSqlQuotes(input));
        }
    }
}
