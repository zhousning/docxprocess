using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Docx.src.docxprocess
{
    class PageSetting
    {
        public void Margins(DocX document, float marginTop, float marginBottom, float marginLeft, float marginRight)
        {

            // Set the page width to be smaller.
            //document.PageWidth = 350f;

            // Set the document margins.
            document.MarginTop = marginTop;
            document.MarginBottom = marginBottom;
            document.MarginLeft = marginLeft;
            document.MarginRight = marginRight;
        }

        public void pageSize(DocX document, float pageWidth, float pageHeight)
        {
            document.PageWidth = pageWidth;
            document.PageHeight = pageHeight;
            document.PageLayout.Orientation = Orientation.Portrait;
        }

        public void pageOrientation(DocX document, string orientation)
        {
            if (orientation == ConstData.PORTRAIT)
            {
                document.PageLayout.Orientation = Orientation.Portrait;
            }else if (orientation == ConstData.LANDSCAPE)
            {
                document.PageLayout.Orientation = Orientation.Landscape;
            }
            
        }
    }
}
