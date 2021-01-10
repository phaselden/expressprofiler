using System;
using System.Text;
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

            if (sql.IndexOf("exec sp_executesql ") > -1)
            {
                var sb = new StringBuilder();
                var parts = StringParser.Split(sql, ' ', '\'');
                var items = StringParser.Split(parts[2], ',', '\'');
                for (var i = 1; i < items.Count; i++)
                {
                    var param = items[i].Length < 80 
                        ? $"-- {items[i]}" 
                        : $"-- {items[i].Substring(0, 80)}...";
                    sb.AppendLine(param.Replace(Environment.NewLine, "  "));
                }
                sb.AppendLine("--");
                sb.AppendLine("-- SQL extracted from sp_executesql and formatted");
                sb.AppendLine("--");
                sb.AppendLine(RemoveSqlQuotes(items[0].Replace("''", "'")));
                sql = sb.ToString();
            }
            
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

        public static string RemoveSqlQuotes(string s)
        {
            s = s.Trim();
            if (s.StartsWith("N"))
                s = s.Substring(1);
            return s.Trim('\'');
        }
    }
}
