using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class TextReplaceService
    {
        private Dictionary<string, string> _replacePatterns;
        public void TextReplaceSet(DocX document, Dictionary<string, string> lists)
        {
            if (lists.Count > 0)
            {
                string source = "(";
                foreach (string key in lists.Keys)
                {
                    source += key + "|";
                }
                source = source.Substring(0, source.Length - 1);
                source += ")";
                this._replacePatterns = lists;
                if (document.FindUniqueByPattern(@source, RegexOptions.IgnoreCase).Count > 0)
                {
                    document.ReplaceText(source, ReplaceTextHandler);
                }
            }
        }

        private string ReplaceTextHandler(string findStr)
        {
            if (_replacePatterns.ContainsKey(findStr))
            {
                return _replacePatterns[findStr];
            }
            return findStr;
        }

    }
}
