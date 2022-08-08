using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class TestService
    {
        public void test()
        {
            using (var document = DocX.Load(@"C:\Users\周宁\Desktop\物联网GIS森林防火智能预警系统.docx"))
            {
                try {
                    document.DifferentOddAndEvenPages = false;
                    document.AddFooters();
                clearFooter(document);
                document.SaveAs(@"C:\Users\周宁\Desktop\新建文件夹\物联网GIS森林防火智能预警系统.docx");
                }
                catch(Exception ex)
                {
                    Debug.WriteLine(ex.Message);
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
            evenFooter.Sections.Clear();
            Debug.WriteLine(evenFooter.ToString());
            Footer oddFooter = document.Footers.Odd;
            oddFooter.Sections.Clear();
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


    }
}
