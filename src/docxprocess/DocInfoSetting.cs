using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Docx.src.model;
using Xceed.Words.NET;

namespace Docx.src.docxprocess
{
    class DocInfoSetting
    {
        public void addCoreProperties(DocX document, DocInfoOption option)
        {
            if (!string.IsNullOrWhiteSpace(option.Subject))
            {
                document.AddCoreProperty("subject", option.Subject);
            }
            if (!string.IsNullOrWhiteSpace(option.Title))
            {
                document.AddCoreProperty("title", option.Title);
            }
            if (!string.IsNullOrWhiteSpace(option.Creator))
            {
                document.AddCoreProperty("creator", option.Creator);
            }
            if (!string.IsNullOrWhiteSpace(option.Description))
            {
                document.AddCoreProperty("description", option.Description);
            }
            if (!string.IsNullOrWhiteSpace(option.Category))
            {
                document.AddCoreProperty("category", option.Category);
            }
            if (!string.IsNullOrWhiteSpace(option.Version))
            {
                document.AddCoreProperty("version", option.Version);
            }


            /*
             * document.AddCoreProperty("created", option.Created);
            document.AddCoreProperty("modified", option.Modified);
             * document.AddCoreProperty("keywords", option.Keywords);
            document.AddCoreProperty("lastModifiedBy", option.LastModifiedBy);
            document.AddCoreProperty("revision", option.Revision);*/
        }
    }
}
