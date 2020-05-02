using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace Docx.src.model
{
    class ParagraphOption:BaseOption
    {
        private string indentationSpecial;
        private float indentationSpecialVal;
        private float indentationBefore;
        private float indentationAfter;
        private float spacing;
        private float spacingAfter;
        private float spacingBefore;
        private float spacingLine;
        private string alignmentString;
        private Alignment alignment;
        private Boolean bold;
        private Boolean italic;
        private Border border;
        private CapsStyle capsStyle;
        private Color color;
        private System.Drawing.Font font;
        private Highlight highlight;       
        private StrikeThrough strikeThrough;
        private Color underlineColor;
        private UnderlineStyle underlineStyle;

        public ParagraphOption(string indentationSpecial, float indentationSpecialVal, float indentationBefore, float indentationAfter, float spacing, float spacingAfter, float spacingBefore, float spacingLine, string alignmentString)
        {
            this.IndentationSpecial = indentationSpecial;
            this.IndentationSpecialVal = indentationSpecialVal;
            this.IndentationBefore = indentationBefore;
            this.IndentationAfter = indentationAfter;
            this.Spacing = spacing;
            this.SpacingAfter = spacingAfter;
            this.SpacingBefore = spacingBefore;
            this.SpacingLine = spacingLine;
            this.alignmentString = alignmentString;
            this.Alignment = setAlignment(alignmentString);
        }

        public string IndentationSpecial { get => indentationSpecial; set => indentationSpecial = value; }
        public float IndentationSpecialVal { get => indentationSpecialVal; set => indentationSpecialVal = value; }
        public float IndentationBefore { get => indentationBefore; set => indentationBefore = value; }
        public float IndentationAfter { get => indentationAfter; set => indentationAfter = value; }
        public float Spacing { get => spacing; set => spacing = value; }
        public float SpacingAfter { get => spacingAfter; set => spacingAfter = value; }
        public float SpacingBefore { get => spacingBefore; set => spacingBefore = value; }
        public float SpacingLine { get => spacingLine; set => spacingLine = value; }
        public string AlignmentString { get => alignmentString; set => alignmentString = value; }
        public Alignment Alignment { get => alignment; set => alignment = value; }
        public bool Bold { get => bold; set => bold = value; }
        public bool Italic { get => italic; set => italic = value; }
        public Border Border { get => border; set => border = value; }
        public CapsStyle CapsStyle { get => capsStyle; set => capsStyle = value; }
        public Color Color { get => color; set => color = value; }
        public System.Drawing.Font Font { get => font; set => font = value; }
        public Highlight Highlight { get => highlight; set => highlight = value; }
        public StrikeThrough StrikeThrough { get => strikeThrough; set => strikeThrough = value; }
        public Color UnderlineColor { get => underlineColor; set => underlineColor = value; }
        public UnderlineStyle UnderlineStyle { get => underlineStyle; set => underlineStyle = value; }
    }
}
