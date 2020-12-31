using System;
using PoorMansTSqlFormatterLib;
using PoorMansTSqlFormatterLib.Formatters;
using SqlVisualizer;

namespace ExpressProfiler
{
    class SqlFormatter
    {
        internal static string Format(string sql)
        {
            if (String.IsNullOrEmpty(sql))
                return "";

            var formatter = GetFormatter(null);
            var wrapper = new HtmlWrapper2(formatter);
            var fullFormatter = new SqlFormattingManager(wrapper);
            var html = fullFormatter.Format(sql);
            return html;
        }

        private static TSqlStandardFormatter GetFormatter(string configString)
        {
            //defaults are as per the object, except disabling colorized/htmlified output

            /*var options = new TSqlStandardFormatterOptions(configString);
            options.HTMLColoring = true;
            options.ExpandCommaLists = false;
            options.ExpandInLists = false;
            options.IndentString = "    "; // The default is "/t", but that's too big in IE and IE doesn't support tab-size CSS
            */
            //return new TSqlStandardFormatter();

            return new TSqlStandardFormatter("    ", 4, 999,
                false, false, false,
                true, true, true,
                false, true,
                true, false);
        }
    }
}
