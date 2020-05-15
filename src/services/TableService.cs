using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
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
                var pTitle = tableDocx.InsertParagraph("当前文档有： " + tables.Count + "个表格\r\n");
                int success = 0;
                int fail = 0;
                for (int i = 0; i < tables.Count; i++)
                {
                    try
                    {
                        var p = tableDocx.InsertParagraph();
                        p.SpacingAfter(20d);
                        p.InsertTableAfterSelf(tables[i]);
                        success++;
                    }
                    catch (Exception ex)
                    {
                        fail++;
                    }                   
                }
                pTitle.Append("提取成功： "+ success + "个表格\r\n");
                pTitle.Append("提取失败： " + fail + "个表格\r\n");
                tableDocx.Save();
            }
        }
    }
}
