using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace Docx.src.model
{
    class HeaderFooterOption:BaseOption
    {
        private System.Drawing.Font font;
        private Color color;
        private string alignmentString;
        private string fontName;
        private float fontSize;
        private Boolean bold;
        private Boolean italic;
        private Alignment alignment;
        private UnderlineStyle underlineStyle;
        private StrikeThrough strikeThrough;
        private string pageText;
        private string firstText;
        private string oddText;
        private string evenText;
        private string image;
        private string pageNumber;
        private Boolean headerFooterLine;



        public HeaderFooterOption(System.Drawing.Font font, Color color, string alignmentString, string pageText, string firstText, string oddText, string evenText, string image, string pageNumber,Boolean headerFooterLine) : this(font, color, alignmentString)
        {
            this.pageText = pageText;
            this.firstText = firstText;
            this.oddText = oddText;
            this.evenText = evenText;
            this.image = image;
            this.pageNumber= pageNumber;
            this.HeaderFooterLine = headerFooterLine;
        }

        public HeaderFooterOption(System.Drawing.Font font, Color color, string alignmentString)
        {
            this.color = color;
            this.FontName = font.Name;
            this.FontSize = font.Size;
            this.bold = font.Bold;
            this.italic = font.Italic;
            this.underlineStyle = font.Underline ? UnderlineStyle.singleLine : UnderlineStyle.none;
            this.strikeThrough = font.Strikeout ? StrikeThrough.strike : StrikeThrough.none;
            this.alignment = setAlignment(alignmentString);
        }

        public string FontName { get => fontName; set => fontName = value; }
        public float FontSize { get => fontSize; set => fontSize = value; }
        public bool Bold { get => bold; set => bold = value; }
        public bool Italic { get => italic; set => italic = value; }
        public Alignment Alignment { get => alignment; set => alignment = value; }
        public UnderlineStyle UnderlineStyle { get => underlineStyle; set => underlineStyle = value; }
        public StrikeThrough StrikeThrough { get => strikeThrough; set => strikeThrough = value; }
        public Color Color { get => color; set => color = value; }
        public string PageText { get => pageText; set => pageText = value; }
        public string FirstText { get => firstText; set => firstText = value; }
        public string OddText { get => oddText; set => oddText = value; }
        public string EvenText { get => evenText; set => evenText = value; }
        public string Image { get => image; set => image = value; }
        public string PageNumber { get => pageNumber; set => pageNumber = value; }
        public bool HeaderFooterLine { get => headerFooterLine; set => headerFooterLine = value; }
    }
}
