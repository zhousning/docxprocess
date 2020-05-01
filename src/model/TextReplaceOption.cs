using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Docx.src.model
{
    class TextReplaceOption
    {
        private string sourceText;
        private string targetText;

        public string SourceText { get => sourceText; set => sourceText = value; }
        public string TargetText { get => targetText; set => targetText = value; }
    }
}
