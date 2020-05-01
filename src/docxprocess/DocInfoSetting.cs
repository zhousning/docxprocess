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
            if (!string.IsNullOrEmpty(option.Subject))
            {
                document.AddCoreProperty("subject", option.Subject);
            }
            if (!string.IsNullOrEmpty(option.Title))
            {
                document.AddCoreProperty("title", option.Title);
            }
            if (!string.IsNullOrEmpty(option.Creator))
            {
                document.AddCoreProperty("creator", option.Creator);
            }
            if (!string.IsNullOrEmpty(option.Description))
            {
                document.AddCoreProperty("description", option.Description);
            }
            if (!string.IsNullOrEmpty(option.Category))
            {
                document.AddCoreProperty("category", option.Category);
            }
            if (!string.IsNullOrEmpty(option.Version))
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
