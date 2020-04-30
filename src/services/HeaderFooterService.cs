using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Docx.src.docxprocess;
using Docx.src.model;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class HeaderFooterService
    {
        private HeaderFooterSetting headerFooterSetting;
        public HeaderFooterService()
        {
            this.headerFooterSetting = new HeaderFooterSetting();
        }


        internal void addHeaders(DocX document, HeaderFooterOption headerOption, bool firstOption, bool oddEvenOption)
        {
            document.AddHeaders();//会将以前有的页眉给删掉
            if (firstOption)
            {
                document.DifferentFirstPage = true;
                this.headerFooterSetting.firstHeader(document, headerOption);
            }
            else
            {
                document.DifferentFirstPage = false;
            }
            if (oddEvenOption)
            {
                document.DifferentOddAndEvenPages = true;
                this.headerFooterSetting.oddHeader(document, headerOption);
                this.headerFooterSetting.evenHeader(document, headerOption);
            }
            else
            {
                this.headerFooterSetting.pageHeader(document, headerOption);
            }
        }

        internal void addFooters(DocX document, HeaderFooterOption footerOption, bool firstOption, bool oddEvenOption)
        {
            document.AddFooters();

            if (firstOption)
            {
                document.DifferentFirstPage = true;
                this.headerFooterSetting.firstFooter(document, footerOption);
            }
            else
            {
                document.DifferentFirstPage = false;
            }
            if (oddEvenOption)
            {
                document.DifferentOddAndEvenPages = true;
                this.headerFooterSetting.oddFooter(document, footerOption);
                this.headerFooterSetting.evenFooter(document, footerOption);
            }
            else
            {
                this.headerFooterSetting.pageFooter(document, footerOption);
            }
        }

       
        public void firstHeader(DocX document, string header, HeaderFooterOption headerOption)
        {
            this.headerFooterSetting.firstHeader(document, headerOption);
        }

        public void firstFooter(DocX document, string footer, HeaderFooterOption footerOption)
        {
            this.headerFooterSetting.firstFooter(document, footerOption);
        }

        public void oddHeader(DocX document, string header, HeaderFooterOption headerOption)
        {
            this.headerFooterSetting.oddHeader(document, headerOption);
        }

        public void evenHeader(DocX document, string header, HeaderFooterOption headerOption)
        {
            this.headerFooterSetting.evenHeader(document,headerOption);
        }

        public void oddFooter(DocX document, string footer, HeaderFooterOption footerOption)
        {
            this.headerFooterSetting.oddFooter(document, footerOption);
        }
        public void evenFooter(DocX document, string footer, HeaderFooterOption footerOption)
        {
            this.headerFooterSetting.evenFooter(document, footerOption);
        }

        public void clearHeader(DocX document)
        {
            this.headerFooterSetting.clearHeader(document);
        }

        public void clearFooter(DocX document)
        {
            this.headerFooterSetting.clearFooter(document);
        }
    }
}
