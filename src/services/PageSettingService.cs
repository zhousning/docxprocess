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

       
    }
}
