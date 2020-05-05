using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class TextReplaceService : BaseService
    {
        private Dictionary<string, string> _replaceTextPatterns;
        private Dictionary<string, DocumentElement> _replaceLinkPatterns;


        public void TextReplaceSet(DocX document, DataGridView ReplaceTextGridView)
        {
            Dictionary<string, string> lists = getTextList(ReplaceTextGridView);

            if (lists.Count > 0)
            {
                string source = "(";
                foreach (string key in lists.Keys)
                {
                    source += key + "|";
                }
                source = source.Substring(0, source.Length - 1);
                source += ")";
                this._replaceTextPatterns = lists;
                if (document.FindUniqueByPattern(@source, RegexOptions.None).Count > 0)
                {
                    document.ReplaceText(source, ReplaceTextHandler);
                }
            }
        }

        public void HyperLinkReplaceSet(DocX document, DataGridView ReplaceLinkGridView)
        {
            foreach (DataGridViewRow row in ReplaceLinkGridView.Rows)
            {
                string source = row.Cells[0].Value == null ? "" : row.Cells[0].Value.ToString();
                Uri target = row.Cells[1].Value == null ? null : GetUri(row.Cells[1].Value.ToString().Trim());
                if (string.IsNullOrWhiteSpace(source) || target == null)
                {
                    continue;
                }
                if (document.FindUniqueByPattern(@source, RegexOptions.None).Count > 0)
                {
                    var linkBlock = document.AddHyperlink(@source, target);
                    
                    document.ReplaceTextWithObject(@source, linkBlock);
                }
            }

        }



        private Uri GetUri(string path)
        {
            //string protocol = @"^(https?|ftp|file|ws)://";
            string pattern = @"^(https?|ftp|file|ws)://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?$";
            Boolean flag = IsMatch(pattern, path);
            if (flag)
            {
                return new Uri(path);
            }
            return new Uri(path);
        }

        private Dictionary<string, string> getTextList(DataGridView ReplaceTextGridView)
        {
            Dictionary<string, string> lists = new Dictionary<string, string>();

            foreach (DataGridViewRow row in ReplaceTextGridView.Rows)
            {
                string source = row.Cells[0].Value == null ? "" : row.Cells[0].Value.ToString();
                string target = row.Cells[1].Value == null ? "" : row.Cells[1].Value.ToString();
                if (source.Length == 0 && target.Length == 0)
                {
                    continue;
                }
                lists.Add(source, target);
            }
            return lists;
        }

        private string ReplaceTextHandler(string findStr)
        {
            if (_replaceTextPatterns.ContainsKey(findStr))
            {
                return _replaceTextPatterns[findStr];
            }
            return findStr;
        }

      
    }
}
