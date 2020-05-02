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
            using (var document = DocX.Create(@"C:\Users\周宁\Desktop\新建文件夹 (2)\SimpleFormattedParagraphs.docx"))
            {
                document.SetDefaultFont(new Xceed.Document.NET.Font("Arial"), 15d, Color.Green);
                document.PageBackground = Color.LightGray;
                document.PageBorders = new Borders(new Border(Xceed.Document.NET.BorderStyle.Tcbs_double, BorderSize.five, 20, Color.Blue));

                // Add a title
                document.InsertParagraph("Formatted paragraphs").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

                // Insert a Paragraph into this document.
                var p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("IndentationHanging This is a simple formatted red bold paragraph")
                .Font(new Xceed.Document.NET.Font("Arial"))
                .FontSize(25)
                .Color(Color.Red)
                .Bold()
                .Append(" containing a blue italic text.").Font(new Xceed.Document.NET.Font("Times New Roman")).Color(Color.Blue).Italic()
              
                .Spacing(30)//字符间距，对append的内容进行设置
                .SpacingLine(20);//多倍行距 10是0.83，对段落整体设置
                p.IndentationHanging = 2f;
                
                /*p.IndentationFirstLine = 2f;
                p.IndentationBefore = 3f;
                p.IndentationAfter = 4f;
                p.LineSpacing = 5f;
                p.LineSpacingAfter = 6f;
                p.LineSpacingBefore = 7f;*/

                // Insert another Paragraph into this document.
                var p2 = document.InsertParagraph();

                // Append some text and add formatting.
                p2.Append("IndentationFirstLine This is a formatted paragraph using spacing, line spacing, ")
                .Font(new Xceed.Document.NET.Font("Courier New"))
                .FontSize(10)
                .Italic()
                .Spacing(5)
                .Append("highlight").Highlight(Highlight.yellow).UnderlineColor(Color.Blue).CapsStyle(CapsStyle.caps)
                .Append(" and strike through.").StrikeThrough(StrikeThrough.strike)
                .IndentationFirstLine = 1.0f;
                p2.LineSpacingBefore = 5f;
                p2.LineSpacingAfter = 15f;
                p2.LineSpacing = 10f;

                // Insert another Paragraph into this document.
                var p3 = document.InsertParagraph();

                // Append some text with 2 TabStopPositions.
                p3.InsertTabStopPosition(Alignment.center, 216f, TabStopPositionLeader.dot)
                .InsertTabStopPosition(Alignment.right, 432f, TabStopPositionLeader.dot)
                .Append("IndentationAfter Text with TabStopPositions on Left\tMiddle\tand Right")
                .FontSize(11d)
                .SpacingAfter(40)
                .IndentationAfter = 1f;
                p3.LineSpacing = 1f;

                // Insert another Paragraph into this document.
                var p4 = document.InsertParagraph();
                p4.Append("IndentationBefore This document is using an Arial green default font of size 15. It's also using a double blue page borders and light gray page background.")
                  .SpacingAfter(40)
                .IndentationBefore = 3f;

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: SimpleFormattedParagraphs.docx\n");
            }
        }

       
    }
}
