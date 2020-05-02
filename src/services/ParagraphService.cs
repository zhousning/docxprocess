using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Docx.src.model;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class ParagraphService
    {


        public void Set(DocX document, TextBox SpaceBefore, TextBox SpaceAfter, TextBox SpaceLineVal, TextBox IndentationSpecialVal, TextBox IndentationBefore, TextBox IndentationAfter, TextBox TextSpace, ComboBox ParagraphAlign, ComboBox SpaceLineType, ComboBox IndentationSpecial)
        {
            string alignmentString = ParagraphAlign.Text;

            float spacingBefore = string.IsNullOrWhiteSpace(SpaceBefore.Text) ? ConstData.FLOAT_INIT : float.Parse(SpaceBefore.Text);
            float spacingAfter = string.IsNullOrWhiteSpace(SpaceAfter.Text) ? ConstData.FLOAT_INIT : float.Parse(SpaceAfter.Text);

            string spaceLineType = SpaceLineType.Text;
            float spacingLine = string.IsNullOrWhiteSpace(SpaceLineVal.Text) ? ConstData.FLOAT_INIT : float.Parse(SpaceLineVal.Text);

            string indentationSpecial = IndentationSpecial.Text;
            float indentationSpecialVal = string.IsNullOrWhiteSpace(IndentationSpecialVal.Text) ? ConstData.FLOAT_INIT : float.Parse(IndentationSpecialVal.Text);

            float indentationBefore = string.IsNullOrWhiteSpace(IndentationBefore.Text) ? ConstData.FLOAT_INIT : float.Parse(IndentationBefore.Text);
            float indentationAfter = string.IsNullOrWhiteSpace(IndentationAfter.Text) ? ConstData.FLOAT_INIT : float.Parse(IndentationAfter.Text);

            float spacing = string.IsNullOrWhiteSpace(TextSpace.Text) ? ConstData.FLOAT_INIT : float.Parse(TextSpace.Text);

            ParagraphOption paragraphOption = new ParagraphOption(indentationSpecial, indentationSpecialVal, indentationBefore, indentationAfter, spacing, spacingAfter, spacingBefore, spacingLine, alignmentString);

            var sections = document.GetSections();
            for (int i = 0; i < sections.Count; ++i)
            {
                var section = sections[i];
                var paragraphs = section.SectionParagraphs;
                for (int j = 0; j < paragraphs.Count; j++)
                {
                    var p = paragraphs[j];
                    GenericSet(p, paragraphOption);
                    IndentationSet(p, paragraphOption);
                    SpaceSet(p, paragraphOption);
                }
            }
            
        }

        public void GenericSet(Paragraph p, ParagraphOption option)
        {
            if(option.AlignmentString != ConstData.NOT_SET)
            {
                p.Alignment = option.Alignment;
            }
        }

        public void IndentationSet(Paragraph p, ParagraphOption option)
        {
            float indentBefore = option.IndentationBefore;
            float indentAfter = option.IndentationAfter;
            string indentSpec = option.IndentationSpecial;
            float indentSpecVal = option.IndentationSpecialVal;
            if (indentBefore != ConstData.FLOAT_INIT)
            {
                p.IndentationBefore = indentBefore;
            }
            if (indentAfter != ConstData.FLOAT_INIT)
            {
                p.IndentationAfter = indentAfter;
            }
            switch (option.IndentationSpecial)
            {
                case ConstData.INDENT_FIRST:
                    p.IndentationFirstLine = option.IndentationSpecialVal;
                    break;
                case ConstData.INDENT_HANG:
                    p.IndentationHanging = option.IndentationSpecialVal;
                    break;
            }
        }
            

        public void SpaceSet(Paragraph p, ParagraphOption option)
        {
            float spaceBefore = option.SpacingBefore;
            float spaceAfter = option.SpacingAfter;
            float spaceLine = option.SpacingLine;
            if (spaceBefore != ConstData.FLOAT_INIT){
                p.SpacingBefore(spaceBefore);
            }
            if (spaceAfter != ConstData.FLOAT_INIT)
            {
                p.SpacingAfter(spaceAfter);
            }
            if (spaceLine != ConstData.FLOAT_INIT)
            {
                p.SpacingLine(spaceLine);
            }
        }
    }
}
