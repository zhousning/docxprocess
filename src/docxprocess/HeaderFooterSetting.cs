using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Docx.src.model;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Docx.src.docxprocess
{
    class HeaderFooterSetting

    {
        public void pageHeader(DocX document, HeaderFooterOption option)
        {
            Header oddHeader = document.Headers.Odd;
            Paragraph oddHeaderP = headerFirstParagraph(oddHeader);
            Header evenHeader = document.Headers.Even;
            Paragraph evenHeaderP = headerFirstParagraph(evenHeader);


            headerSet(document, evenHeaderP, option.PageText, option);
            headerSet(document, oddHeaderP, option.PageText, option);
        }

        public void firstHeader(DocX document, HeaderFooterOption option)
        {
            Header firstHeader = document.Headers.First;
            Paragraph firstHeaderP = headerFirstParagraph(firstHeader);

            headerSet(document, firstHeaderP, option.FirstText, option);
        }

        public void oddHeader(DocX document, HeaderFooterOption option)
        {
            Header oddHeader = document.Headers.Odd;
            Paragraph oddHeaderP = headerFirstParagraph(oddHeader);

            headerSet(document, oddHeaderP, option.OddText, option);
            
        }

        public void evenHeader(DocX document, HeaderFooterOption option)
        {
            Header evenHeader = document.Headers.Even;
            Paragraph evenHeaderP = headerFirstParagraph(evenHeader);

            headerSet(document, evenHeaderP, option.EvenText, option);
        }

        public void pageFooter(DocX document, HeaderFooterOption option)
        {
            Footer oddFooter = document.Footers.Odd;
            Paragraph oddFooterP = footerFirstParagraph(oddFooter);
            Footer evenFooter = document.Footers.Even;
            Paragraph evenFooterP = footerFirstParagraph(evenFooter);
            
            footerSet(document, evenFooterP, option.PageText, option);
            footerSet(document, oddFooterP, option.PageText, option);
        }
        public void firstFooter(DocX document, HeaderFooterOption option)
        {
            Footer firstFooter = document.Footers.First;
            Paragraph firstFooterP = footerFirstParagraph(firstFooter);

            footerSet(document, firstFooterP, option.FirstText, option);
        }

        public void oddFooter(DocX document, HeaderFooterOption option)
        {
            Footer oddFooter = document.Footers.Odd;
            Paragraph oddFooterP = footerFirstParagraph(oddFooter);

            footerSet(document, oddFooterP, option.OddText, option);
        }

        public void evenFooter(DocX document, HeaderFooterOption option)
        {
            Footer evenFooter = document.Footers.Even;
            Paragraph evenFooterP = footerFirstParagraph(evenFooter);

            footerSet(document, evenFooterP, option.EvenText, option);
        }


        private void headerSet(DocX document, Paragraph paragraph, string title, HeaderFooterOption option)
        {
            PAppendImage(document, paragraph, option);
            headerLineOption(paragraph, option);
            Options(paragraph.Append(title), option);
        }

        private void footerSet(DocX document, Paragraph paragraph, string title, HeaderFooterOption option)
        {
            pageNumber(paragraph, option);//pagenumber 要放在image之前，不然图片会被插入两次
            PAppendImage(document, paragraph, option);         
            footerLineOption(paragraph, option);
            Options(paragraph.Append(title), option);
        }

        private void pageNumber(Paragraph paragraph, HeaderFooterOption option)
        {
            switch (option.PageNumber)
            {
                case ConstData.PAGENUMBER1:
                    Options(paragraph.AppendPageNumber(PageNumberFormat.normal), option);
                    break;
                case ConstData.PAGENUMBER2:
                    Options(paragraph.Append(" / "), option);
                    paragraph.InsertPageNumber(PageNumberFormat.normal, 0);
                    paragraph.InsertPageCount(PageNumberFormat.normal, 2);
                    break;
                case ConstData.PAGENUMBER3:
                    Options(paragraph.Append("第页"), option);
                    paragraph.InsertPageNumber(PageNumberFormat.normal, 1);
                    break;
                case ConstData.PAGENUMBER4:
                    Options(paragraph.Append("第页 共页"), option);
                    paragraph.InsertPageNumber(PageNumberFormat.normal, 1);
                    paragraph.InsertPageCount(PageNumberFormat.normal, 5);
                    break;
                case ConstData.PAGENUMBER5:
                    Options(paragraph.AppendPageNumber(PageNumberFormat.roman), option);
                    break;
                case ConstData.PAGENUMBER6:
                    Options(paragraph.Append(" / "), option);
                    paragraph.InsertPageNumber(PageNumberFormat.roman, 0);
                    paragraph.InsertPageCount(PageNumberFormat.roman, 2);
                    break;
                case ConstData.PAGENUMBER7:
                    Options(paragraph.Append("第页"), option);
                    paragraph.InsertPageNumber(PageNumberFormat.roman, 1);
                    break;
                case ConstData.PAGENUMBER8:
                    Options(paragraph.Append("第页 共页"), option);
                    paragraph.InsertPageNumber(PageNumberFormat.roman, 1);
                    paragraph.InsertPageCount(PageNumberFormat.roman, 5);
                    break;

            }
                      
        }


        public void clearHeader(DocX document)
        {
            Headers headers = document.Headers;
            if (document.DifferentFirstPage)
            {
                Header firstHeader = document.Headers.First;
                
                if (firstHeader != null)
                {
                    int firstCount = firstHeader.Paragraphs.Count;
                    for (int i = 0; i < firstCount; i++)
                    {
                        firstHeader.RemoveParagraphAt(i);
                    }
                }
            }
            Header evenHeader = document.Headers.Even;
            Header oddHeader = document.Headers.Odd;
            if (evenHeader != null)
            {
                int pCount = evenHeader.Paragraphs.Count;
                for (int i = 0; i < pCount; i++)
                {
                    evenHeader.RemoveParagraphAt(i);
                }
                int tCount = evenHeader.Tables.Count;
                for (int j=0; j<tCount; j++)
                {
                    evenHeader.Tables[j].Remove();
                }
                int iCount = evenHeader.Images.Count;
                for(int k=0; k<iCount; k++)
                {
                    evenHeader.Images.RemoveAt(k);
                }
                int picCount = evenHeader.Images.Count;
                for (int n = 0; n < picCount; n++)
                {
                    evenHeader.Pictures.RemoveAt(n);
                }
            }

            if (oddHeader != null)
            {
                int pCount = oddHeader.Paragraphs.Count;
                for (int i = 0; i < pCount; i++)
                {
                    oddHeader.RemoveParagraphAt(i);
                }
                int tCount = oddHeader.Tables.Count;
                for (int j = 0; j < tCount; j++)
                {
                    oddHeader.Tables[j].Remove();
                }
                int iCount = oddHeader.Images.Count;
                for (int k = 0; k < iCount; k++)
                {
                    oddHeader.Images.RemoveAt(k);
                }
                int picCount = oddHeader.Images.Count;
                for (int n = 0; n < picCount; n++)
                {
                    oddHeader.Pictures.RemoveAt(n);
                }
            }
        }

        public void clearFooter(DocX document)
        {
            if (document.DifferentFirstPage)
            {
                Footer firstFooter = document.Footers.First;
                if (firstFooter != null)
                {
                    int firstCount = firstFooter.Paragraphs.Count;
                    for (int i = 0; i < firstCount; i++)
                    {
                        firstFooter.RemoveParagraphAt(i);
                    }
                }
            }
            Footer evenFooter = document.Footers.Even;           
            Footer oddFooter = document.Footers.Odd;
            if (evenFooter != null)
            {
                int pCount = evenFooter.Paragraphs.Count;
                for (int i = 0; i < pCount; i++)
                {
                    evenFooter.RemoveParagraphAt(i);
                }
                int tCount = evenFooter.Tables.Count;
                for (int j = 0; j < tCount; j++)
                {
                    evenFooter.Tables[j].Remove();
                }
                int iCount = evenFooter.Images.Count;
                for (int k = 0; k < iCount; k++)
                {
                    evenFooter.Images.RemoveAt(k);
                }
                int picCount = evenFooter.Images.Count;
                for (int n = 0; n < picCount; n++)
                {
                    evenFooter.Pictures.RemoveAt(n);
                }
            }

            if (oddFooter != null)
            {
                int pCount = oddFooter.Paragraphs.Count;
                for (int i = 0; i < pCount; i++)
                {
                    oddFooter.RemoveParagraphAt(i);
                }
                int tCount = oddFooter.Tables.Count;
                for (int j = 0; j < tCount; j++)
                {
                    oddFooter.Tables[j].Remove();
                }
                int iCount = oddFooter.Images.Count;
                for (int k = 0; k < iCount; k++)
                {
                    oddFooter.Images.RemoveAt(k);
                }
                int picCount = oddFooter.Images.Count;
                for (int n = 0; n < picCount; n++)
                {
                    oddFooter.Pictures.RemoveAt(n);
                }
            }
        }


        private void PAppendImage(DocX document, Paragraph paragraph, HeaderFooterOption option)
        {
            if (!string.IsNullOrWhiteSpace(option.Image))
            {
                paragraph.AppendPicture(document.AddImage(@option.Image).CreatePicture());
            }
        }

        private void Options(Paragraph paragraph, HeaderFooterOption option)
        {
            paragraph.Font(option.FontName).FontSize(option.FontSize).Bold(option.Bold).Italic(option.Italic).StrikeThrough(option.StrikeThrough).UnderlineStyle(option.UnderlineStyle).Color(option.Color).Alignment = option.Alignment;
        }

        private void headerLineOption(Paragraph paragraph, HeaderFooterOption option)
        {
            if (option.HeaderFooterLine)
            {
                paragraph.InsertHorizontalLine(HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single);
            }

        }

        private void footerLineOption(Paragraph paragraph, HeaderFooterOption option)
        {
            if (option.HeaderFooterLine)
            {
                paragraph.InsertHorizontalLine(HorizontalBorderPosition.top, BorderStyle.Tcbs_single);
            }
        }
            

        private Paragraph headerFirstParagraph(Header header)
        {
            Paragraph headerP = header.Paragraphs.Count > 0 ? header.Paragraphs.First() : header.InsertParagraph();
            return headerP;
        }

        private Paragraph footerFirstParagraph(Footer footer)
        {
            Paragraph footerP = footer.Paragraphs.Count > 0 ? footer.Paragraphs.First() : footer.InsertParagraph();
            return footerP;
        }
    }
}
