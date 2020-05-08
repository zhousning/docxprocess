using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Docx.src;
using Docx.src.controllers;
using Docx.src.model;
using Docx.src.workers;
using Xceed.Words.NET;

namespace Docx
{
    public partial class MainForm : Form
    {
        private MainController mainController;
        private BackgroundWorker bgWorker;

        public MainForm()
        {
            InitializeComponent();
            InitializeData();
        }

        private void InitializeData()
        {
            MainFormOption mainFormOption = new MainFormOption(outPutFolder, ExtractImageCheckBox, ExtractHyperLinkCheckBox, ExtractTable, ReplaceLinkGridView, notSetMargin, notSetPageSize, topMargin, bottomMargin, leftMargin, rightMargin, pageWidth, pageHeight, pageSetOrientation, clearHeader, clearFooter, firstHeaderFooter, oddEvenHeaderFooter,
             notSetHeader, notSetFooter, headerFontDialog, headerAlignComBox, headerColorDialog,
             pageHeader, firstHeader, oddHeader, evenHeader, headerImagePath, headerLine, footerFontDialog, footerAlignComBox, footerColorDialog,
             pageFooter, firstFooter, oddFooter, evenFooter, footerImagePath, footerLine, pageNumberComBox, DocTitle, DocSubject, DocCategory,
             DocDescription, DocCreator, DocVersion, DocEditPrctRemove,
             DocEditPrctRemove, DocEditPassword, TaskProcessBtn, PdfExportBtn, OutputFolderBtn, inputFolderBtn, StopWork, fileGrid, toolStripProgressBar, todoTask,
             ReplaceTextGridView, CreateTimeCheckBox, DocCreateTime, UpdateTimeCheckBox, DocUpdateTime);
            this.mainController = new MainController();
            BackWorker backWorker = new BackWorker(mainFormOption);
            bgWorker = backWorker.getWorker();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.fileGrid.Columns[0].FillWeight = 140;
            this.fileGrid.Columns[1].FillWeight = 80;

            pageSetOrientation.SelectedIndex = 0;
            headerAlignComBox.SelectedIndex = 1;
            footerAlignComBox.SelectedIndex = 1;
            //(*.jpg,*.png,*.jpeg,*.bmp,*.gif)|*.jgp;*.png;*.jpeg;*.bmp;*.gif|All files(*.*)|*.*
            headerImageDialog.Filter = "(*.jpg,*.png,*.jpeg)|*.jpg;*.png;*.jpeg";
            footerImageDialog.Filter = "(*.jpg,*.png,*.jpeg)|*.jpg;*.png;*.jpeg";

            PageNumberComBox_Load();
        }

        private void Button2_Click_1(object sender, EventArgs e)
        {
            mainController.OutputFolderBtnEvent(outPutFolder);
        }

        private void NotSetMargin_CheckedChanged(object sender, EventArgs e)
        {
            mainController.NotSetMarginEvent(notSetMargin, topMargin,
            bottomMargin, leftMargin, rightMargin);
        }

        private void PageAddToTask_CheckedChanged(object sender, EventArgs e)
        {
            mainController.addToTaskCheck(todoTask, pageAddToTask);
        }

        private void TabPage1_Click(object sender, EventArgs e)
        {

        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }


        private void SplitContainer1_Panel2_Paint_1(object sender, PaintEventArgs e)
        {

        }


        private void SplitContainer2_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void SplitContainer2_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void TabPage2_Click(object sender, EventArgs e)
        {

        }

        private void FileGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void FileGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string result = fileGrid.Rows[e.RowIndex].Cells["result"].Value.ToString();

                if (result == ConstData.FAIL)
                {
                    fileGrid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    fileGrid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                }
                else if (result == ConstData.WITHOUT_FILE)
                {
                    fileGrid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Orange;
                    fileGrid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                }
            }
        }

        private void GroupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void NotSetPageSize_CheckedChanged(object sender, EventArgs e)
        {
            mainController.NotSetPageSizeEvent(notSetPageSize, pageWidth, pageHeight);

        }


        private void headerFooterToTask_CheckedChanged(object sender, EventArgs e)
        {
            mainController.addToTaskCheck(todoTask, headerFooterToTask);
        }

        private void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (firstHeaderFooter.Checked)
            {
                firstHeader.Enabled = true;
                firstFooter.Enabled = true;
            }
            else
            {
                firstHeader.Enabled = false;
                firstFooter.Enabled = false;
            }
        }

        private void OddEvenHeaderFooter_CheckedChanged(object sender, EventArgs e)
        {
            if (oddEvenHeaderFooter.Checked)
            {
                pageHeader.Enabled = false;
                pageFooter.Enabled = false;
                oddHeader.Enabled = true;
                oddFooter.Enabled = true;
                evenHeader.Enabled = true;
                evenFooter.Enabled = true;
            }
            else
            {
                pageHeader.Enabled = true;
                pageFooter.Enabled = true;
                oddHeader.Enabled = false;
                oddFooter.Enabled = false;
                evenHeader.Enabled = false;
                evenFooter.Enabled = false;
            }
        }

        

        private void HeaderFontBtn_Click(object sender, EventArgs e)
        {
            headerFontDialog.ShowDialog();
        }

        private void FooterFontBtn_Click(object sender, EventArgs e)
        {
            footerFontDialog.ShowDialog();
        }

        private void HeaderColorBtn_Click(object sender, EventArgs e)
        {
            headerColorDialog.ShowDialog();
        }

        private void FooterColorBtn_Click(object sender, EventArgs e)
        {
            footerColorDialog.ShowDialog();
        }

        private void HeaderImageBtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = headerImageDialog.ShowDialog();
            string filename = headerImageDialog.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(filename))
            {
                headerImagePath.Text = filename;
            }
        }

        private void FooterImageBtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = footerImageDialog.ShowDialog();
            string filename = footerImageDialog.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(filename))
            {
                footerImagePath.Text = filename;
            }
        }


        private void ClearHeaderImage_Click(object sender, EventArgs e)
        {
            headerImagePath.Text = "";
        }

        private void ClearFooterImageBtn_Click(object sender, EventArgs e)
        {
            footerImagePath.Text = "";
        }

        private void PageNumberComBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void PageNumberComBox_Load()
        {
            pageNumberComBox.Items.Add(ConstData.PAGENUMBERNO);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER1);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER2);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER3);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER4);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER5);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER6);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER7);
            pageNumberComBox.Items.Add(ConstData.PAGENUMBER8);
            pageNumberComBox.SelectedIndex = 0;
        }



        private void TabPage1_Click_1(object sender, EventArgs e)
        {

        }


        private void HeaderGroupBox_Enter(object sender, EventArgs e)
        {

        }


        private void Label20_Click(object sender, EventArgs e)
        {

        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {

        }


        private void GroupBox6_Enter_1(object sender, EventArgs e)
        {

        }

        private void Title_TextChanged(object sender, EventArgs e)
        {

        }

        private void Subject_TextChanged(object sender, EventArgs e)
        {

        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void PageInfoToTask_CheckedChanged(object sender, EventArgs e)
        {
            mainController.addToTaskCheck(todoTask, pageInfoToTask);
        }

        private void DocVersion_TextChanged(object sender, EventArgs e)
        {

        }


        private void DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }


        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void TextBox1_TextChanged_2(object sender, EventArgs e)
        {

        }

        private void Label21_Click(object sender, EventArgs e)
        {

        }

        private void Tabcontrol_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void DocPrct_CheckedChanged(object sender, EventArgs e)
        {

        }


        private void TextReplaceTab_Click(object sender, EventArgs e)
        {

        }

        private void TextReplacetoTask_CheckedChanged(object sender, EventArgs e)
        {
            mainController.addToTaskCheck(todoTask, textReplacetoTask);
        }

        private void GroupBox10_Enter(object sender, EventArgs e)
        {

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void TextBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Label24_Click(object sender, EventArgs e)
        {

        }

        private void TabPage1_Click_2(object sender, EventArgs e)
        {

        }

        private void GroupBox10_Enter_1(object sender, EventArgs e)
        {

        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void ParagraphCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            mainController.addToTaskCheck(todoTask, ParagraphToTask);
        }

        private void IndentationSpecial_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void SpaceBefore_TextChanged(object sender, EventArgs e)
        {

        }

        private void ImageToTask_CheckedChanged(object sender, EventArgs e)
        {
            mainController.addToTaskCheck(todoTask, ImageToTask);
        }

        private void ExtractTable_CheckedChanged(object sender, EventArgs e)
        {

        }


        private void InputFolderBtn_Click_1(object sender, EventArgs e)
        {
            mainController.InputFolderBtnEvent(inputFolder, fileGrid);
        }

        private void startChangeBtn()
        {
            TaskProcessBtn.Enabled = false;
            PdfExportBtn.Enabled = false;
            OutputFolderBtn.Enabled = false;
            inputFolderBtn.Enabled = false;
            StopWork.Enabled = true;
            fileGrid.AllowUserToDeleteRows = false;
        }

      
        private void TaskProcessBtn_Click(object sender, EventArgs e)
        {
            if (todoTask.Items.Count == 0)
            {
                MessageBox.Show("当前没有待处理任务", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (fileGrid.Rows.Count == 0)
            {
                MessageBox.Show("请添加文件", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            startChangeBtn();
            string title = TaskProcessBtn.Text;
            this.bgWorker.RunWorkerAsync(title);
        }

        private void PdfExportBtn_Click_1(object sender, EventArgs e)
        {
            if (fileGrid.Rows.Count == 0)
            {
                MessageBox.Show("请添加文件", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            startChangeBtn();
            string title = PdfExportBtn.Text;
            this.bgWorker.RunWorkerAsync(title);
        }

        private void StopWork_Click_1(object sender, EventArgs e)
        {
            this.bgWorker.CancelAsync();
        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void ToolStripProgressBar1_Click(object sender, EventArgs e)
        {

        }



        private void ToolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        /* todoTask listbox 多选右键删除，暂时先不用
        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int count = todoTask.SelectedItems.Count;
            List<string> itemValues = new List<string>();
            if (count != 0)
            {
                for (int i = 0; i < count; i++)
                {
                    itemValues.Add(todoTask.SelectedItems[i].ToString());
                }
                foreach (string item in itemValues)
                {
                    todoTask.Items.Remove(item);
                }
            }
        }*/
    }
}
