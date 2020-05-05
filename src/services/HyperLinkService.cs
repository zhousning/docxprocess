using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class HyperLinkService : BaseService
    {
        public void extractHyperLink(DocX document, string output)
        {
            string outputPath = makedir(output);
            var hyperlinks = document.Hyperlinks;
            using (StreamWriter sw = new StreamWriter(outputPath + "超链接.txt"))
            {
                string link = "当前文档有" + hyperlinks.Count +"个超链接\r\n";
                for (int i = 0; i < hyperlinks.Count; i++)
                {
                    link += hyperlinks[i].Text + ": " + hyperlinks[i].Uri + "\r\n";                    
                }
                sw.WriteLine(link);
            }
        }

        public void replaceHyperLink(DocX document)
        {
            var hyperlinks = document.Hyperlinks;
            var h3 = document.AddHyperlink("hyperlink", new Uri("http://www.xceed.com/"));
            // Add a paragraph.
            var p3 = document.InsertParagraph("An hyperlink pointing to a bookmark of this Document has been added at the end of this paragraph: ");
            p3.ReplaceTextWithObject("hyperlink", h3);
        }

    }
}
