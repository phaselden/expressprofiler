using System;
using PoorMansTSqlFormatterLib;
using PoorMansTSqlFormatterLib.Formatters;
using SqlVisualizer;

namespace ExpressProfiler
{
    public class SqlFormatter
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

        public static string FindSpExecuteSql(string sql)
        {
            const string start = "exec sp_executesql N'";
            var startIndex = sql.IndexOf(start, StringComparison.OrdinalIgnoreCase);
            if (startIndex == -1)
                return null;
            var sqlStartIndex = startIndex + start.Length;
            var endIndex = GetEndOfStringIndex(sql, sqlStartIndex);
            if (endIndex == -1)
                return "ERROR";
            return UnescapeSingleQuotes(sql.Substring(sqlStartIndex, endIndex - sqlStartIndex));
        }

        private static string UnescapeSingleQuotes(string s)
        {
            return s.Replace("''", "'");
        }

        private static int GetEndOfStringIndex(string s, int startIndex)
        {
            const char quote = '\'';
            var inEscape = false;

            for (var i = startIndex; i < s.Length; i++)
            {
                if (s[i] != quote) 
                    continue;

                if (!inEscape)
                {
                    if (PeekChar(s, i + 1) != quote)
                        return i;
                    else
                        inEscape = true;
                }
                else
                {
                    inEscape = false;
                }
            }

            return -1;
        }

        private static char PeekChar(string sql, int index)
        {
            if (index < sql.Length)
                return sql[index];
            return default(char);
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
