using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class TableService : BaseService
    {
        public void extractTable(DocX document, string output)
        {
            string outputPath = makedir(output);
            var tables = document.Tables;

            using(var tableDocx = DocX.Create(outputPath + "表格.docx"))
            {
                tableDocx.InsertParagraph("当前文档有" + tables.Count + "个表格");
                for (int i = 0; i < tables.Count; i++)
                {
                    var p = tableDocx.InsertParagraph();
                    p.SpacingAfter(20d);
                    p.InsertTableAfterSelf(tables[i]);
                }
                tableDocx.Save();
            }
        }
    }
}
