using System;
using System.Collections.Generic;

namespace ExpressProfiler
{
    public class StringParser
    {
        public static List<string> Split(string s, char separator, char quoteChar)
        {
            var result = new List<string>();

            var startIndex = 0;
            var currentIndex = 0;
            var inQuotedString = false;

            while (currentIndex < s.Length)
            {
                if (s[currentIndex] == quoteChar)
                {
                    inQuotedString = !inQuotedString;
                }
                // Split if a separator is found, unless inside a quoted string
                else if (s[currentIndex] == separator && !inQuotedString)
                {
                    var address = s.Substring(startIndex, currentIndex - startIndex);
                    result.Add(address);
                    startIndex = currentIndex + 1;
                }
                currentIndex++;
            }

            if (currentIndex > startIndex)
            {
                var address = s.Substring(startIndex, currentIndex - startIndex);
                result.Add(address);
            }

            if (inQuotedString)
                throw new FormatException("Unclosed quote");

            return result;
        }

        /*public List<Span<string>> SplitIntoSpans(string s, char separator, char quoteChar)
        {
            var result = new List<string>();

            var startIndex = 0;
            var currentIndex = 0;
            var inQuotedString = false;

            while (currentIndex < s.Length)
            {
                if (s[currentIndex] == quoteChar)
                {
                    inQuotedString = !inQuotedString;
                }
                // Split if a comma is found, unless inside a quoted string
                else if (s[currentIndex] == separator && !inQuotedString)
                {
                    var address = s.Substring(startIndex, currentIndex - startIndex);
                    result.Add(address);
                    startIndex = currentIndex + 1;
                }
                currentIndex++;
            }

            if (currentIndex > startIndex)
            {
                var address = s.Substring(startIndex, currentIndex - startIndex);
                result.Add(address);
            }

            if (inQuotedString)
                throw new FormatException("Unclosed quote");

            return result;
        }*/

    }
}