using System;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Docx.src.docxprocess;
using Xceed.Words.NET;
using System.Drawing;
using Docx.src.model;
using Xceed.Document.NET;
using System.IO;
using System.Drawing.Imaging;

namespace Docx.src.services
{
    class PageSettingService
    {
        private PageSetting pageSetting;

        public PageSettingService()
        {
            this.pageSetting = new PageSetting();
        }
        public void marginSetting(DocX document, Boolean notSetMargin, string topMargin, string bottomMargin, string leftMargin, string rightMargin)
        {
            if (!notSetMargin)
            {
                float marginTop = float.Parse(topMargin) * ConstData.POUND;
                float marginBottom = float.Parse(bottomMargin) * ConstData.POUND;
                float marginLeft = float.Parse(leftMargin) * ConstData.POUND;
                float marginRight = float.Parse(rightMargin) * ConstData.POUND;
                this.pageSetting.Margins(document, marginTop, marginBottom, marginLeft, marginRight);
            }
        }

        public void pageSizeSetting(DocX document, Boolean notSetPageSize, string pageWidth, string pageHeight)
        {
            if (!notSetPageSize)
            {
                float width = float.Parse(pageWidth) * ConstData.POUND;
                float height = float.Parse(pageHeight) * ConstData.POUND;
                this.pageSetting.pageSize(document, width, height);
            }
        }

        public void pageOrientation(DocX document, string orientation)
        {
            this.pageSetting.pageOrientation(document, orientation);
        }

        public void test()
        {
            using (var document = DocX.Create(@"C:\Users\周宁\Desktop\新建文件夹 (3)\最终版 24版.docx"))
            {
                var h3 = document.AddHyperlink("hyperlink", new Uri("http://www.baidu.com"));
                // Add a paragraph.
                var p3 = document.InsertParagraph("An hyperlink pointing to a bookmark of this Document has been added at the end of this paragraph: ").Bold().FontSize(30d);
                //document.ReplaceTextWithObject("hyperlink", h3);
                string style = p3.StyleName;
                p3.AppendHyperlink(h3);
                document.Save();
                Console.WriteLine("\tCreated: Mod ifyImage.docx\n");
            }
        }

       
    }
}
