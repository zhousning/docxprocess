using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Docx.src.model
{
    class FormValOption
    {
        private string processTitle;
        private string outPutFolder;
        private Boolean extractImageCheckBox;
        private Boolean extractHyperLinkCheckBox;
        private Boolean extractTable;
        private Boolean notSetMargin;
        private Boolean notSetPageSize;
        private string topMargin;
        private string bottomMargin;
        private string leftMargin;
        private string rightMargin;
        private string pageWidth;
        private string pageHeight;
        private string pageSetOrientation;
        private Boolean clearHeader;
        private Boolean clearFooter;
        private Boolean firstHeaderFooter;
        private Boolean oddEvenHeaderFooter;
        private Boolean notSetHeader;
        private Boolean notSetFooter;
        private Font headerFontDialog;
        private string headerAlignComBox;
        private Color headerColorDialog;
        private string pageHeader;
        private string firstHeader;
        private string oddHeader;
        private string evenHeader;
        private string headerImagePath;
        private Boolean headerLine;
        private Font footerFontDialog;
        private string footerAlignComBox;
        private Color footerColorDialog;
        private string pageFooter;
        private string firstFooter;
        private string oddFooter;
        private string evenFooter;
        private string footerImagePath;
        private Boolean footerLine;
        private string pageNumberComBox;
        private string docTitle;
        private string docSubject;
        private string docCategory;
        private string docDescription;
        private string docCreator;
        private string docVersion;
        private Boolean docEditPrctCheckBox;
        private Boolean docEditPrctRemove;
        private string docEditPassword;
        private Dictionary<string, string> fileGrid;
        private List<string> todoTask;
        private Dictionary<string, string> replaceTextGridView;
        private Dictionary<string, string> replaceLinkGridView;
        private Boolean createTimeCheckBox;
        private DateTime docCreateTime;
        private Boolean updateTimeCheckBox;
        private DateTime docUpdateTime;

        public FormValOption(string outPutFolder, bool extractImageCheckBox, bool extractHyperLinkCheckBox, bool extractTable, bool notSetMargin, bool notSetPageSize, string topMargin, string bottomMargin, string leftMargin, string rightMargin, string pageWidth, string pageHeight, string pageSetOrientation, bool clearHeader, bool clearFooter, bool firstHeaderFooter, bool oddEvenHeaderFooter, bool notSetHeader, bool notSetFooter, Font headerFontDialog, string headerAlignComBox, Color headerColorDialog, string pageHeader, string firstHeader, string oddHeader, string evenHeader, string headerImagePath, bool headerLine, Font footerFontDialog, string footerAlignComBox, Color footerColorDialog, string pageFooter, string firstFooter, string oddFooter, string evenFooter, string footerImagePath, bool footerLine, string pageNumberComBox, string docTitle, string docSubject, string docCategory, string docDescription, string docCreator, string docVersion, bool docEditPrctCheckBox, bool docEditPrctRemove, string docEditPassword, Dictionary<string,string> fileGrid, List<string> todoTask, Dictionary<string, string> replaceTextGridView, Dictionary<string, string> replaceLinkGridView, bool createTimeCheckBox, DateTime docCreateTime, bool updateTimeCheckBox, DateTime docUpdateTime)
        {
            this.OutPutFolder = outPutFolder;
            this.ExtractImageCheckBox = extractImageCheckBox;
            this.ExtractHyperLinkCheckBox = extractHyperLinkCheckBox;
            this.ExtractTable = extractTable;
            this.NotSetMargin = notSetMargin;
            this.NotSetPageSize = notSetPageSize;
            this.TopMargin = topMargin;
            this.BottomMargin = bottomMargin;
            this.LeftMargin = leftMargin;
            this.RightMargin = rightMargin;
            this.PageWidth = pageWidth;
            this.PageHeight = pageHeight;
            this.PageSetOrientation = pageSetOrientation;
            this.ClearHeader = clearHeader;
            this.ClearFooter = clearFooter;
            this.FirstHeaderFooter = firstHeaderFooter;
            this.OddEvenHeaderFooter = oddEvenHeaderFooter;
            this.NotSetHeader = notSetHeader;
            this.NotSetFooter = notSetFooter;
            this.HeaderFontDialog = headerFontDialog;
            this.HeaderAlignComBox = headerAlignComBox;
            this.HeaderColorDialog = headerColorDialog;
            this.PageHeader = pageHeader;
            this.FirstHeader = firstHeader;
            this.OddHeader = oddHeader;
            this.EvenHeader = evenHeader;
            this.HeaderImagePath = headerImagePath;
            this.HeaderLine = headerLine;
            this.FooterFontDialog = footerFontDialog;
            this.FooterAlignComBox = footerAlignComBox;
            this.FooterColorDialog = footerColorDialog;
            this.PageFooter = pageFooter;
            this.FirstFooter = firstFooter;
            this.OddFooter = oddFooter;
            this.EvenFooter = evenFooter;
            this.FooterImagePath = footerImagePath;
            this.FooterLine = footerLine;
            this.PageNumberComBox = pageNumberComBox;
            this.DocTitle = docTitle;
            this.DocSubject = docSubject;
            this.DocCategory = docCategory;
            this.DocDescription = docDescription;
            this.DocCreator = docCreator;
            this.DocVersion = docVersion;
            this.DocEditPrctCheckBox = docEditPrctCheckBox;
            this.DocEditPrctRemove = docEditPrctRemove;
            this.DocEditPassword = docEditPassword;
            this.FileGrid = fileGrid;
            this.TodoTask = todoTask;
            this.ReplaceTextGridView = replaceTextGridView;
            this.ReplaceLinkGridView = replaceLinkGridView;
            this.CreateTimeCheckBox = createTimeCheckBox;
            this.DocCreateTime = docCreateTime;
            this.UpdateTimeCheckBox = updateTimeCheckBox;
            this.DocUpdateTime = docUpdateTime;
        }

        public string OutPutFolder { get => outPutFolder; set => outPutFolder = value; }
        public bool ExtractImageCheckBox { get => extractImageCheckBox; set => extractImageCheckBox = value; }
        public bool ExtractHyperLinkCheckBox { get => extractHyperLinkCheckBox; set => extractHyperLinkCheckBox = value; }
        public bool ExtractTable { get => extractTable; set => extractTable = value; }
        public bool NotSetMargin { get => notSetMargin; set => notSetMargin = value; }
        public bool NotSetPageSize { get => notSetPageSize; set => notSetPageSize = value; }
        public string TopMargin { get => topMargin; set => topMargin = value; }
        public string BottomMargin { get => bottomMargin; set => bottomMargin = value; }
        public string LeftMargin { get => leftMargin; set => leftMargin = value; }
        public string RightMargin { get => rightMargin; set => rightMargin = value; }
        public string PageWidth { get => pageWidth; set => pageWidth = value; }
        public string PageHeight { get => pageHeight; set => pageHeight = value; }
        public string PageSetOrientation { get => pageSetOrientation; set => pageSetOrientation = value; }
        public bool ClearHeader { get => clearHeader; set => clearHeader = value; }
        public bool ClearFooter { get => clearFooter; set => clearFooter = value; }
        public bool FirstHeaderFooter { get => firstHeaderFooter; set => firstHeaderFooter = value; }
        public bool OddEvenHeaderFooter { get => oddEvenHeaderFooter; set => oddEvenHeaderFooter = value; }
        public bool NotSetHeader { get => notSetHeader; set => notSetHeader = value; }
        public bool NotSetFooter { get => notSetFooter; set => notSetFooter = value; }
        public Font HeaderFontDialog { get => headerFontDialog; set => headerFontDialog = value; }
        public string HeaderAlignComBox { get => headerAlignComBox; set => headerAlignComBox = value; }
        public Color HeaderColorDialog { get => headerColorDialog; set => headerColorDialog = value; }
        public string PageHeader { get => pageHeader; set => pageHeader = value; }
        public string FirstHeader { get => firstHeader; set => firstHeader = value; }
        public string OddHeader { get => oddHeader; set => oddHeader = value; }
        public string EvenHeader { get => evenHeader; set => evenHeader = value; }
        public string HeaderImagePath { get => headerImagePath; set => headerImagePath = value; }
        public bool HeaderLine { get => headerLine; set => headerLine = value; }
        public Font FooterFontDialog { get => footerFontDialog; set => footerFontDialog = value; }
        public string FooterAlignComBox { get => footerAlignComBox; set => footerAlignComBox = value; }
        public Color FooterColorDialog { get => footerColorDialog; set => footerColorDialog = value; }
        public string PageFooter { get => pageFooter; set => pageFooter = value; }
        public string FirstFooter { get => firstFooter; set => firstFooter = value; }
        public string OddFooter { get => oddFooter; set => oddFooter = value; }
        public string EvenFooter { get => evenFooter; set => evenFooter = value; }
        public string FooterImagePath { get => footerImagePath; set => footerImagePath = value; }
        public bool FooterLine { get => footerLine; set => footerLine = value; }
        public string PageNumberComBox { get => pageNumberComBox; set => pageNumberComBox = value; }
        public string DocTitle { get => docTitle; set => docTitle = value; }
        public string DocSubject { get => docSubject; set => docSubject = value; }
        public string DocCategory { get => docCategory; set => docCategory = value; }
        public string DocDescription { get => docDescription; set => docDescription = value; }
        public string DocCreator { get => docCreator; set => docCreator = value; }
        public string DocVersion { get => docVersion; set => docVersion = value; }
        public bool DocEditPrctCheckBox { get => docEditPrctCheckBox; set => docEditPrctCheckBox = value; }
        public bool DocEditPrctRemove { get => docEditPrctRemove; set => docEditPrctRemove = value; }
        public string DocEditPassword { get => docEditPassword; set => docEditPassword = value; }
        public List<string> TodoTask { get => todoTask; set => todoTask = value; }
        public Dictionary<string, string> ReplaceTextGridView { get => replaceTextGridView; set => replaceTextGridView = value; }
        public Dictionary<string, string> ReplaceLinkGridView { get => replaceLinkGridView; set => replaceLinkGridView = value; }
        public bool CreateTimeCheckBox { get => createTimeCheckBox; set => createTimeCheckBox = value; }
        public DateTime DocCreateTime { get => docCreateTime; set => docCreateTime = value; }
        public bool UpdateTimeCheckBox { get => updateTimeCheckBox; set => updateTimeCheckBox = value; }
        public DateTime DocUpdateTime { get => docUpdateTime; set => docUpdateTime = value; }
        public Dictionary<string, string> FileGrid { get => fileGrid; set => fileGrid = value; }
        public string ProcessTitle { get => processTitle; set => processTitle = value; }
    }
}
