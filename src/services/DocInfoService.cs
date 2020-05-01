using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Docx.src.docxprocess;
using Docx.src.model;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class DocInfoService
    {
        private DocInfoSetting docInfoSetting;
        public DocInfoService()
        {
            this.docInfoSetting = new DocInfoSetting();
        }

        public void addCoreProperties(DocX document, DocInfoOption option)
        {
            this.docInfoSetting.addCoreProperties(document, option);
        }

        public void UpdateFileTime(string targetFile, Boolean createTime, DateTime createTimeVal, Boolean updateTime, DateTime updateTimeVal)
        {
            if (createTime)
            {
                File.SetCreationTime(targetFile, createTimeVal);
            }
            if (updateTime)
            {
                File.SetLastWriteTime(targetFile, updateTimeVal);
            }
        }

        public void DocProtect(DocX document, Boolean editProtect, string password)
        {
            if (editProtect)
            {
                document.RemoveProtection();
                document.AddPasswordProtection(EditRestrictions.readOnly, password);
            }
        }
        public void DocRemoveProtect(DocX document, Boolean removeProtect)
        {
            if (removeProtect)
            {
                document.RemoveProtection();
            }
        }
    }
}
