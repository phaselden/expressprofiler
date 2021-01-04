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

            var subSql = FindSpExecuteSql(sql);
            if (subSql != null)
            {
                sql += Environment.NewLine + Environment.NewLine + 
                       "-- SQL Extracted and Formatted" 
                       + Environment.NewLine + Environment.NewLine + subSql;
            }
            
            var formatter = GetFormatter(null);
            var wrapper = new HtmlWrapper2(formatter);
            var fullFormatter = new SqlFormattingManager(wrapper);
            var html = fullFormatter.Format(sql);
            return html;
        }

        private static string FindSpExecuteSql(string sql)
        {
            const string start = "exec sp_executesql N'";
            var startIndex = sql.IndexOf(start, StringComparison.OrdinalIgnoreCase);
            if (startIndex == -1)
                return null;
            var sqlStartIndex = startIndex + start.Length;
            var endIndex = sql.IndexOf("'", sqlStartIndex, StringComparison.Ordinal);
            if (endIndex == -1)
                return "ERROR";
            return sql.Substring(sqlStartIndex, endIndex - sqlStartIndex);
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
