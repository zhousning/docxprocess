using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Docx.src.model
{
    class MainFormOption
    {
        private TextBox outPutFolder;
        private CheckBox extractImageCheckBox;
        private CheckBox extractHyperLinkCheckBox;
        private CheckBox extractTable;
        private DataGridView replaceLinkGridView;
        private CheckBox notSetMargin;
        private CheckBox notSetPageSize;
        private NumericUpDown topMargin;
        private NumericUpDown bottomMargin;
        private NumericUpDown leftMargin;
        private NumericUpDown rightMargin;
        private NumericUpDown pageWidth;
        private NumericUpDown pageHeight;
        private ComboBox pageSetOrientation;
        private CheckBox clearHeader;
        private CheckBox clearFooter;
        private CheckBox firstHeaderFooter;
        private CheckBox oddEvenHeaderFooter;
        private CheckBox notSetHeader;
        private CheckBox notSetFooter;
        private FontDialog headerFontDialog;
        private ComboBox headerAlignComBox;
        private ColorDialog headerColorDialog;
        private TextBox pageHeader;
        private TextBox firstHeader;
        private TextBox oddHeader;
        private TextBox evenHeader;
        private TextBox headerImagePath;
        private CheckBox headerLine;
        private FontDialog footerFontDialog;
        private ComboBox footerAlignComBox;
        private ColorDialog footerColorDialog;
        private TextBox pageFooter;
        private TextBox firstFooter;
        private TextBox oddFooter;
        private TextBox evenFooter;
        private TextBox footerImagePath;
        private CheckBox footerLine;
        private ComboBox pageNumberComBox;
        private TextBox docTitle;
        private TextBox docSubject;
        private TextBox docCategory;
        private TextBox docDescription;
        private TextBox docCreator;
        private TextBox docVersion;
        private CheckBox docEditPrctCheckBox;
        private CheckBox docEditPrctRemove;
        private TextBox docEditPassword;
        private Button taskProcessBtn;
        private Button pdfExportBtn;
        private Button outputFolderBtn;
        private Button inputFolderBtn;
        private Button stopWork;
        private DataGridView fileGrid;
        private ToolStripProgressBar toolStripProgressBar;
        private ListBox todoTask;
        private DataGridView replaceTextGridView;
        private CheckBox createTimeCheckBox;
        private DateTimePicker docCreateTime;
        private CheckBox updateTimeCheckBox;
        private DateTimePicker docUpdateTime;

        public MainFormOption(TextBox outPutFolder, CheckBox extractImageCheckBox, CheckBox extractHyperLinkCheckBox, CheckBox extractTable, DataGridView replaceLinkGridView, CheckBox notSetMargin, CheckBox notSetPageSize, NumericUpDown topMargin, NumericUpDown bottomMargin, NumericUpDown leftMargin, NumericUpDown rightMargin, NumericUpDown pageWidth, NumericUpDown pageHeight, ComboBox pageSetOrientation, CheckBox clearHeader, CheckBox clearFooter, CheckBox firstHeaderFooter, CheckBox oddEvenHeaderFooter, CheckBox notSetHeader, CheckBox notSetFooter, FontDialog headerFontDialog, ComboBox headerAlignComBox, ColorDialog headerColorDialog, TextBox pageHeader, TextBox firstHeader, TextBox oddHeader, TextBox evenHeader, TextBox headerImagePath, CheckBox headerLine, FontDialog footerFontDialog, ComboBox footerAlignComBox, ColorDialog footerColorDialog, TextBox pageFooter, TextBox firstFooter, TextBox oddFooter, TextBox evenFooter, TextBox footerImagePath, CheckBox footerLine, ComboBox pageNumberComBox, TextBox docTitle, TextBox docSubject, TextBox docCategory, TextBox docDescription, TextBox docCreator, TextBox docVersion, CheckBox docEditPrctCheckBox, CheckBox docEditPrctRemove, TextBox docEditPassword
            , Button taskProcessBtn, Button pdfExportBtn, Button outputFolderBtn, Button inputFolderBtn, Button stopWork, DataGridView fileGrid, ToolStripProgressBar toolStripProgressBar, ListBox todoTask, DataGridView replaceTextGridView,
            CheckBox createTimeCheckBox,DateTimePicker docCreateTime,CheckBox updateTimeCheckBox,DateTimePicker docUpdateTime)
        {
            this.CreateTimeCheckBox = createTimeCheckBox;
            this.DocCreateTime = docCreateTime;
            this.UpdateTimeCheckBox = updateTimeCheckBox;
            this.DocUpdateTime = docUpdateTime;
            this.OutPutFolder = outPutFolder;
            this.ExtractImageCheckBox = extractImageCheckBox;
            this.ExtractHyperLinkCheckBox = extractHyperLinkCheckBox;
            this.ExtractTable = extractTable;
            this.ReplaceLinkGridView = replaceLinkGridView;
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
            this.TaskProcessBtn = taskProcessBtn;
            this.PdfExportBtn = pdfExportBtn;
            this.OutputFolderBtn = outputFolderBtn;
            this.InputFolderBtn = inputFolderBtn;
            this.StopWork = stopWork;
            this.FileGrid = fileGrid;
            this.ToolStripProgressBar = toolStripProgressBar;
            this.TodoTask = todoTask;
            this.ReplaceTextGridView = replaceTextGridView;
        }

        public TextBox OutPutFolder { get => outPutFolder; set => outPutFolder = value; }
        public CheckBox ExtractImageCheckBox { get => extractImageCheckBox; set => extractImageCheckBox = value; }
        public CheckBox ExtractHyperLinkCheckBox { get => extractHyperLinkCheckBox; set => extractHyperLinkCheckBox = value; }
        public CheckBox ExtractTable { get => extractTable; set => extractTable = value; }
        public DataGridView ReplaceLinkGridView { get => replaceLinkGridView; set => replaceLinkGridView = value; }
        public CheckBox NotSetMargin { get => notSetMargin; set => notSetMargin = value; }
        public CheckBox NotSetPageSize { get => notSetPageSize; set => notSetPageSize = value; }
        public NumericUpDown TopMargin { get => topMargin; set => topMargin = value; }
        public NumericUpDown BottomMargin { get => bottomMargin; set => bottomMargin = value; }
        public NumericUpDown LeftMargin { get => leftMargin; set => leftMargin = value; }
        public NumericUpDown RightMargin { get => rightMargin; set => rightMargin = value; }
        public NumericUpDown PageWidth { get => pageWidth; set => pageWidth = value; }
        public NumericUpDown PageHeight { get => pageHeight; set => pageHeight = value; }
        public ComboBox PageSetOrientation { get => pageSetOrientation; set => pageSetOrientation = value; }
        public CheckBox ClearHeader { get => clearHeader; set => clearHeader = value; }
        public CheckBox ClearFooter { get => clearFooter; set => clearFooter = value; }
        public CheckBox FirstHeaderFooter { get => firstHeaderFooter; set => firstHeaderFooter = value; }
        public CheckBox OddEvenHeaderFooter { get => oddEvenHeaderFooter; set => oddEvenHeaderFooter = value; }
        public CheckBox NotSetHeader { get => notSetHeader; set => notSetHeader = value; }
        public CheckBox NotSetFooter { get => notSetFooter; set => notSetFooter = value; }
        public FontDialog HeaderFontDialog { get => headerFontDialog; set => headerFontDialog = value; }
        public ComboBox HeaderAlignComBox { get => headerAlignComBox; set => headerAlignComBox = value; }
        public ColorDialog HeaderColorDialog { get => headerColorDialog; set => headerColorDialog = value; }
        public TextBox PageHeader { get => pageHeader; set => pageHeader = value; }
        public TextBox FirstHeader { get => firstHeader; set => firstHeader = value; }
        public TextBox OddHeader { get => oddHeader; set => oddHeader = value; }
        public TextBox EvenHeader { get => evenHeader; set => evenHeader = value; }
        public TextBox HeaderImagePath { get => headerImagePath; set => headerImagePath = value; }
        public CheckBox HeaderLine { get => headerLine; set => headerLine = value; }
        public FontDialog FooterFontDialog { get => footerFontDialog; set => footerFontDialog = value; }
        public ComboBox FooterAlignComBox { get => footerAlignComBox; set => footerAlignComBox = value; }
        public ColorDialog FooterColorDialog { get => footerColorDialog; set => footerColorDialog = value; }
        public TextBox PageFooter { get => pageFooter; set => pageFooter = value; }
        public TextBox FirstFooter { get => firstFooter; set => firstFooter = value; }
        public TextBox OddFooter { get => oddFooter; set => oddFooter = value; }
        public TextBox EvenFooter { get => evenFooter; set => evenFooter = value; }
        public TextBox FooterImagePath { get => footerImagePath; set => footerImagePath = value; }
        public CheckBox FooterLine { get => footerLine; set => footerLine = value; }
        public ComboBox PageNumberComBox { get => pageNumberComBox; set => pageNumberComBox = value; }
        public TextBox DocTitle { get => docTitle; set => docTitle = value; }
        public TextBox DocSubject { get => docSubject; set => docSubject = value; }
        public TextBox DocCategory { get => docCategory; set => docCategory = value; }
        public TextBox DocDescription { get => docDescription; set => docDescription = value; }
        public TextBox DocCreator { get => docCreator; set => docCreator = value; }
        public TextBox DocVersion { get => docVersion; set => docVersion = value; }
        public CheckBox DocEditPrctCheckBox { get => docEditPrctCheckBox; set => docEditPrctCheckBox = value; }
        public CheckBox DocEditPrctRemove { get => docEditPrctRemove; set => docEditPrctRemove = value; }
        public TextBox DocEditPassword { get => docEditPassword; set => docEditPassword = value; }
        public Button TaskProcessBtn { get => taskProcessBtn; set => taskProcessBtn = value; }
        public Button PdfExportBtn { get => pdfExportBtn; set => pdfExportBtn = value; }
        public Button OutputFolderBtn { get => outputFolderBtn; set => outputFolderBtn = value; }
        public Button InputFolderBtn { get => inputFolderBtn; set => inputFolderBtn = value; }
        public Button StopWork { get => stopWork; set => stopWork = value; }
        public DataGridView FileGrid { get => fileGrid; set => fileGrid = value; }
        public ToolStripProgressBar ToolStripProgressBar { get => toolStripProgressBar; set => toolStripProgressBar = value; }
        public ListBox TodoTask { get => todoTask; set => todoTask = value; }
        public DataGridView ReplaceTextGridView { get => replaceTextGridView; set => replaceTextGridView = value; }
        public CheckBox CreateTimeCheckBox { get => createTimeCheckBox; set => createTimeCheckBox = value; }
        public DateTimePicker DocCreateTime { get => docCreateTime; set => docCreateTime = value; }
        public CheckBox UpdateTimeCheckBox { get => updateTimeCheckBox; set => updateTimeCheckBox = value; }
        public DateTimePicker DocUpdateTime { get => docUpdateTime; set => docUpdateTime = value; }
    }
}
