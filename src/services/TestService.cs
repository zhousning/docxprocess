using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class TestService
    {
        public void test()
        {
            using (var document = DocX.Load(@"C:\Users\周宁\Desktop\新建文件夹\设置商品的参数、属性、规格.docx"))
            {
                document.RemoveProtection();
                document.SaveAs(@"C:\Users\周宁\Desktop\新建文件夹 (4)\设置商品的参数、属性、规格.docx");
                Console.WriteLine("\tCreated: Mod ifyImage.docx\n");
            }
        }

    }
}
