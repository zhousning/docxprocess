using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Docx.src;
using Docx.src.docxprocess;
using Docx.src.model;
using Docx.src.services;
using Xceed.Words.NET;

namespace Docx
{
    public partial class MainForm : Form
    {
        private PageSettingService pageSettingService;
        private HeaderFooterService headerFooterService;
        public MainForm()
        {
            InitializeComponent();
            this.pageSettingService = new PageSettingService();
            this.headerFooterService = new HeaderFooterService();
        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                String path = dilog.SelectedPath;
                DirectoryInfo directoryInfo = new DirectoryInfo(path);
                FileInfo[] files = directoryInfo.GetFiles();

                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                dt.Columns.Add("filename", typeof(string));
                dt.Columns.Add("filepath", typeof(string));
                dt.Columns.Add("filesize", typeof(string));
                dt.Columns.Add("result", typeof(string));
                foreach (FileInfo f in files)
                {
                    string filename = f.Name;
                    string filepath = f.FullName;
                    string filesize = System.Math.Ceiling(f.Length / 1024.0) + " KB";

                    string extension = Path.GetExtension(filepath);
                    if (extension == ".docx")
                    {
                        DataRow row = dt.NewRow();
                        row["filename"] = filename;
                        row["filepath"] = filepath;
                        row["filesize"] = filesize;
                        row["result"] = "";
                        dt.Rows.Add(row);
                    }
                }
                ds.Tables.Add(dt);
                fileGrid.DataSource = ds.Tables[0];
            }
        }


        private void MainForm_Load(object sender, EventArgs e)
        {
            this.fileGrid.Columns[0].FillWeight = 140;
            this.fileGrid.Columns[1].FillWeight = 80;

            pageSetOrientation.SelectedIndex = 0;
            headerAlignComBox.SelectedIndex = 1;
            footerAlignComBox.SelectedIndex = 1;
            //(*.jpg,*.png,*.jpeg,*.bmp,*.gif)|*.jgp;*.png;*.jpeg;*.bmp;*.gif|All files(*.*)|*.*
            headerImageDialog.Filter= "(*.jpg,*.png,*.jpeg)|*.jpg;*.png;*.jpeg";
            footerImageDialog.Filter= "(*.jpg,*.png,*.jpeg)|*.jpg;*.png;*.jpeg";

            PageNumberComBox_Load();
        }



        private void Button1_Click(object sender, EventArgs e)
        {
            List<string> tasks = todoTask.Items.Cast<string>().ToList();
            if (tasks.Count == 0)
            {
                MessageBox.Show("当前没有待处理任务", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            startProcess.Enabled = false;
            string outputDirectory = outPutFolder.Text;
            if (string.IsNullOrEmpty(outputDirectory))
            {
                MessageBox.Show("请选择输出目录", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                startProcess.Enabled = true;
                return;
            }
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("filename", typeof(string));
            dt.Columns.Add("filepath", typeof(string));
            dt.Columns.Add("filesize", typeof(string));
            dt.Columns.Add("result", typeof(string));
            foreach (DataGridViewRow row in fileGrid.Rows)
            {
                string filepath = row.Cells["filepath"].Value.ToString();
                string filename = row.Cells["filename"].Value.ToString();
                string filesize = row.Cells["filesize"].Value.ToString();
                string result = "";
                try
                {
                    if (!String.IsNullOrEmpty(filepath) && File.Exists(filepath))
                    {
                        string targetFile = outputDirectory + @"\" + filename;
                        using (var document = DocX.Load(filepath))
                        {
                            foreach (string title in tasks)
                            {
                                if (title == pageSettingTab.Text)
                                {
                                    PageSet(document);
                                }
                                else if (title == headerFooterTab.Text)
                                {
                                    HeaderFooterSet(document);
                                }
                            }
                            document.SaveAs(targetFile);
                            result = ConstData.SUCCESS;
                        }
                    }
                    else
                    {
                        result = ConstData.WITHOUT_FILE;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    result = ConstData.FAIL;
                }
                finally
                {
                    DataRow dataRow = dt.NewRow();
                    dataRow["filename"] = filename;
                    dataRow["filepath"] = filepath;
                    dataRow["filesize"] = filesize;
                    dataRow["result"] = result;
                    dt.Rows.Add(dataRow);
                }
            }
            ds.Tables.Add(dt);
            fileGrid.DataSource = ds.Tables[0];
            startProcess.Enabled = true;
            //MessageBox.Show("处理完毕", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void PageSet(DocX document)
        {
            pageSettingService.marginSetting(document, notSetMargin.Checked, topMargin.Text, bottomMargin.Text, leftMargin.Text, rightMargin.Text);
            pageSettingService.pageSizeSetting(document, notSetPageSize.Checked, pageWidth.Text, pageHeight.Text);
            pageSettingService.pageOrientation(document, pageSetOrientation.Text);
        }

        private void HeaderFooterSet(DocX document)
        {
            if (clearHeader.Checked)
            {
                headerFooterService.clearHeader(document);
            }
            if (clearFooter.Checked)
            {
                headerFooterService.clearFooter(document);
            }

            Boolean firstOption = firstHeaderFooter.Checked;
            Boolean oddEvenOption = oddEvenHeaderFooter.Checked;
            if (!notSetHeader.Checked && !clearHeader.Checked)
            {
                Font headerFont = headerFontDialog.Font;
                string headerAlign = headerAlignComBox.Text;
                Color headerColor = headerColorDialog.Color;
                string pageHeaderText = pageHeader.Text;
                string firstHeaderText = firstHeader.Text;
                string oddHeaderText = oddHeader.Text;
                string evenHeaderText = evenHeader.Text;
                string headerImage = headerImagePath.Text;
                Boolean headerLineBool = headerLine.Checked;
                
                
                HeaderFooterOption headerOption = new HeaderFooterOption(headerFont, headerColor, headerAlign, pageHeaderText, firstHeaderText, oddHeaderText, evenHeaderText, headerImage, "", headerLineBool);

                headerFooterService.addHeaders(document, headerOption, firstOption, oddEvenOption);
            }

            if (!notSetFooter.Checked && !clearFooter.Checked)
            {
                Font footerFont = footerFontDialog.Font;
                string footerAlign = footerAlignComBox.Text;
                Color footerColor = footerColorDialog.Color;
                string pageFooterText = pageFooter.Text;
                string firstFooterText = firstFooter.Text;
                string oddFooterText = oddFooter.Text;
                string evenFooterText = evenFooter.Text;
                string footerImage = footerImagePath.Text;
                string pageNumber = pageNumberComBox.Text;
                Boolean footerLineBool = footerLine.Checked;
                HeaderFooterOption footerOption = new HeaderFooterOption(footerFont, footerColor, footerAlign, pageFooterText, firstFooterText, oddFooterText, evenFooterText, footerImage, pageNumber, footerLineBool);

                headerFooterService.addFooters(document, footerOption, firstOption, oddEvenOption);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                outPutFolder.Text = dilog.SelectedPath;
            }
        }


        private void NotSetMargin_CheckedChanged(object sender, EventArgs e)
        {
            if (notSetMargin.Checked)
            {
                topMargin.Enabled = false;
                bottomMargin.Enabled = false;
                leftMargin.Enabled = false;
                rightMargin.Enabled = false;
            }
            else
            {
                topMargin.Enabled = true;
                bottomMargin.Enabled = true;
                leftMargin.Enabled = true;
                rightMargin.Enabled = true;
            }
        }

        private void PageAddToTask_CheckedChanged(object sender, EventArgs e)
        {
            if (pageAddToTask.Checked)
            {
                todoTask.Items.Add(pageSettingTab.Text);
            }
            else
            {
                todoTask.Items.Remove(pageSettingTab.Text);
            }
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void TabPage1_Click(object sender, EventArgs e)
        {

        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }


        private void TopMargin_TextChanged(object sender, EventArgs e)
        {

        }
        private void SplitContainer1_Panel1_Paint_1(object sender, PaintEventArgs e)
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

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void TodoTask_SelectedIndexChanged(object sender, EventArgs e)
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
            if (notSetPageSize.Checked)
            {
                pageWidth.Enabled = false;
                pageHeight.Enabled = false;
            }
            else
            {
                pageWidth.Enabled = true;
                pageHeight.Enabled = true;
            }
        }



        private void 页面设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new PageSettingForm().Show();
        }



        private void headerFooterToTask_CheckedChanged(object sender, EventArgs e)
        {
            if (headerFooterToTask.Checked)
            {
                todoTask.Items.Add(headerFooterTab.Text);
            }
            else
            {
                todoTask.Items.Remove(headerFooterTab.Text);
            }
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

        private void NotSetHeader_CheckedChanged(object sender, EventArgs e)
        {
            /*if (notSetHeader.Checked)
            {
                foreach (Control col in headerGroupBox.Controls)
                {
                    if (col.Text != "不设置")
                    {
                        col.Enabled = false;
                    }
                }
            }
            else
            {
                if (firstHeaderFooter.Checked)
                {
                    firstHeader.Enabled = true;
                }
                if (oddEvenHeaderFooter.Checked)
                {
                    pageHeader.Enabled = false;
                    oddHeader.Enabled = true;
                    evenHeader.Enabled = true;
                }
                else
                {
                    pageHeader.Enabled = true;
                }
                headerFontBtn.Enabled = true;
                headerAlignComBox.Enabled = true;
                headerColorBtn.Enabled = true;
            }*/
        }

        private void NotSetFooter_CheckedChanged(object sender, EventArgs e)
        {
            /*if (notSetFooter.Checked)
            {
                foreach (Control col in footerGroupBox.Controls)
                {
                    if (col.Text != "不设置")
                    {
                        col.Enabled = false;
                    }
                }
            }
            else
            {
                if (firstHeaderFooter.Checked)
                {
                    firstHeader.Enabled = true;
                }
                if (oddEvenHeaderFooter.Checked)
                {
                    pageFooter.Enabled = false;
                    oddFooter.Enabled = true;
                    evenFooter.Enabled = true;
                }
                else
                {
                    pageFooter.Enabled = true;
                }
                footerFontBtn.Enabled = true;
                footerAlignComBox.Enabled = true;
                footerColorBtn.Enabled = true;
            }*/
        }

        private void HeaderImageBtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = headerImageDialog.ShowDialog();
            string filename = headerImageDialog.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(filename))
            {
                headerImagePath.Text = filename;
            }
        }

        private void FooterImageBtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = footerImageDialog.ShowDialog();
            string filename = footerImageDialog.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(filename))
            {
                footerImagePath.Text = filename;
            }
        }

        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

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

        private void Button1_Click_1(object sender, EventArgs e)
        {
            
                Console.WriteLine("\tInsertHorizontalLine()");

                using (var document = DocX.Create(@"C:\Users\周宁\Desktop\新建文件夹 (2)\InsertHorizontalLine.docx"))
                {
                    // Add a title
                    document.InsertParagraph("Adding top or bottom Horizontal lines").FontSize(15d).SpacingAfter(50d).Alignment = Xceed.Document.NET.Alignment.center;

                    // Add a paragraph with a single line.
                    var p = document.InsertParagraph();
                    p.Append("This is a paragraph with a single bottom line.").Font(new Xceed.Document.NET.Font("Arial")).FontSize(15);
                    p.InsertHorizontalLine(Xceed.Document.NET.HorizontalBorderPosition.bottom, Xceed.Document.NET.BorderStyle.Tcbs_single);
                    p.SpacingAfter(20);

                    // Add a paragraph with a double green line.
                    var p2 = document.InsertParagraph();
                    p2.Append("This is a paragraph with a double bottom colored line.").Font(new Xceed.Document.NET.Font("Arial")).FontSize(15);
                    p2.InsertHorizontalLine(Xceed.Document.NET.HorizontalBorderPosition.bottom, Xceed.Document.NET.BorderStyle.Tcbs_double, 6, 1, Color.Green);
                    p2.SpacingAfter(20);

                    // Add a paragraph with a triple red line.
                    var p3 = document.InsertParagraph();
                    p3.Append("This is a paragraph with a triple bottom colored line.").Font(new Xceed.Document.NET.Font("Arial")).FontSize(15);
                    p3.InsertHorizontalLine(Xceed.Document.NET.HorizontalBorderPosition.bottom, Xceed.Document.NET.BorderStyle.Tcbs_triple, 6, 1, Color.Red);
                    p3.SpacingAfter(20);

                    // Add a paragraph with a single spaced line.
                    var p4 = document.InsertParagraph();
                    p4.Append("This is a paragraph with a spaced bottom line.").Font(new Xceed.Document.NET.Font("Arial")).FontSize(15);
                    p4.InsertHorizontalLine(Xceed.Document.NET.HorizontalBorderPosition.bottom, Xceed.Document.NET.BorderStyle.Tcbs_single, 6, 12);
                    p4.SpacingAfter(20);

                    // Add a paragraph with a single large line.
                    var p5 = document.InsertParagraph();
                    p5.Append("This is a paragraph with a large bottom line.").Font(new Xceed.Document.NET.Font("Arial")).FontSize(15);
                    p5.InsertHorizontalLine(Xceed.Document.NET.HorizontalBorderPosition.bottom, Xceed.Document.NET.BorderStyle.Tcbs_single, 25);
                    p5.SpacingAfter(60);

                    // Add a paragraph with a wave blue top line.
                    var p6 = document.InsertParagraph();
                    p6.Append("This is a paragraph with a wave blue top line.").Font(new Xceed.Document.NET.Font("Arial")).FontSize(15);
                    p6.InsertHorizontalLine(Xceed.Document.NET.HorizontalBorderPosition.top, Xceed.Document.NET.BorderStyle.Tcbs_wave, 6, 1, Color.FromArgb(0, 0, 255));
                    p5.SpacingAfter(20);

                    document.Save();
                    Console.WriteLine("\tCreated: InsertHorizontalLine.docx\n");
                }
            
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
