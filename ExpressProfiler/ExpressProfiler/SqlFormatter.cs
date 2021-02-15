using System;
using System.Collections.Generic;
using System.Text;
using PoorMansTSqlFormatterLib;
using PoorMansTSqlFormatterLib.Formatters;
using SqlVisualizer;

namespace EdtDbProfiler
{
    public class SqlParam
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }

    public class ParseResult
    {
        public ParseResult()
        {
            Parameters = new List<SqlParam>();
        }

        public string Sql { get; set; }
        public List<SqlParam> Parameters { get; }
    }

    public class SqlFormatter
    {
        internal static string Format(string sql)
        {
            if (String.IsNullOrEmpty(sql))
                return "";

            var parseResult = ParseSql(sql);
            
            var formatter = GetFormatter(null);
            var wrapper = new HtmlWrapper2(formatter);
            var fullFormatter = new SqlFormattingManager(wrapper);

            var sb = new StringBuilder();
            if (parseResult.Parameters.Count > 0)
            {
                foreach (var p in parseResult.Parameters)
                {
                    var param = p.Value.Length < 80
                        ? $"-- {p.Value}"
                        : $"-- {p.Value.Substring(0, 80)}...";
                    sb.AppendLine(param.Replace(Environment.NewLine, "  "));
                }
                sb.AppendLine("--");
                sb.AppendLine("-- SQL extracted from sp_executesql and formatted");
                sb.AppendLine("--");
                sb.AppendLine(parseResult.Sql);
                return fullFormatter.Format(sb.ToString());
            }
            return fullFormatter.Format(parseResult.Sql);
        }

        public static ParseResult ParseSql(string sql)
        {
            ParseResult result = new ParseResult();
            result.Sql = sql;

            if (sql.IndexOf("exec sp_executesql ") > -1)
            {
                var sb = new StringBuilder();
                var parts = StringParser.Split(sql, ' ', '\'');
                var items = StringParser.Split(parts[2], ',', '\'');
                for (var i = 1; i < items.Count; i++)
                {
                    var p = new SqlParam();
                    p.Value = items[i].Replace(Environment.NewLine, "  ");
                    result.Parameters.Add(p);
                }

                result.Sql = RemoveSqlQuotes(items[0].Replace("''", "'"));
                result.Sql = sb.ToString();
            }
            
            return result;
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
