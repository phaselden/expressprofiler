//Traceutils assembly
//writen by Locky, 2009. 

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Text;

namespace EdtDbProfiler
{
    class RtfBuilder
    {
        private readonly StringBuilder _stringBuilder = new StringBuilder();
        private readonly List<Color> _colortable = new List<Color>();
        private readonly StringCollection _fontTable = new StringCollection();
        
        private Color _foreColor;
        public Color ForeColor
        {
            set
            {
                if (!_colortable.Contains(value)) { _colortable.Add(value); }
                if (value != _foreColor)
                {
                    _stringBuilder.Append(String.Format("\\cf{0} ", _colortable.IndexOf(value) + 1));
                }
                _foreColor = value;
            }
        }
        
        private Color _backColor;
        public Color BackColor
        {
            set
            {
                if (!_colortable.Contains(value)) { _colortable.Add(value); }
                if (value != _backColor)
                {
                    _stringBuilder.Append(String.Format("\\highlight{0} ", _colortable.IndexOf(value) + 1));
                }
                _backColor = value;
            }
        }
        
        public RtfBuilder()
        {
            ForeColor = Color.FromKnownColor(KnownColor.WindowText);
            BackColor = Color.FromKnownColor(KnownColor.Window);
            _defaultFontSize = 20F;
        }

        public void AppendLine()
        {
            _stringBuilder.AppendLine("\\line");
        }

        public void Append(string value)
        {
            if (!String.IsNullOrEmpty(value))
            {
                value = CheckChar(value);
                if (value.IndexOf(Environment.NewLine) >= 0)
                {
                    var lines = value.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                    foreach (string line in lines)
                    {
                        _stringBuilder.Append(line);
                        _stringBuilder.Append("\\line ");
                    }
                }
                else
                {
                    _stringBuilder.Append(value);
                }
            }
        }

        private static readonly char[] _slashable = new[] { '{', '}', '\\' };
        private readonly float _defaultFontSize;

        private static string CheckChar(string value)
        {
            if (!String.IsNullOrEmpty(value))
            {
                if (value.IndexOfAny(_slashable) >= 0)
                {
                    value = value.Replace("{", "\\{").Replace("}", "\\}").Replace("\\", "\\\\");
                }
                var replaceUni = false;
                for (var i = 0; i < value.Length; i++)
                {
                    if (value[i] > 255)
                    {
                        replaceUni = true;
                        break;
                    }
                }
                if (replaceUni)
                {
                    var sb = new StringBuilder();
                    for (var i = 0; i < value.Length; i++)
                    {
                        if (value[i] <= 255)
                        {
                            sb.Append(value[i]);
                        }
                        else
                        {
                            sb.Append("\\u");
                            sb.Append((int)value[i]);
                            sb.Append("?");
                        }
                    }
                    value = sb.ToString();
                }
            }
            
            return value;
        }

        public new string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang3081");
            sb.Append("{\\fonttbl");
            for (var i = 0; i < _fontTable.Count; i++)
            {
                try
                {
                    sb.Append(string.Format(_fontTable[i], i));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            sb.AppendLine("}");
            sb.Append("{\\colortbl ;");
            foreach (var color in _colortable)
            {
                sb.AppendFormat("\\red{0}\\green{1}\\blue{2};", color.R, color.G, color.B);
            }
            sb.AppendLine("}");
            sb.Append("\\viewkind4\\uc1\\pard\\plain\\f0");
            sb.AppendFormat("\\fs{0} ", _defaultFontSize);
            sb.AppendLine();
            sb.Append(_stringBuilder.ToString());
            sb.Append("}");
            return sb.ToString();
        }
    }
}