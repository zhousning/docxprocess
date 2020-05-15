using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Docx.src.model;
using Docx.src.services;
using Xceed.Words.NET;

namespace Docx.src.controllers
{
    class MainController
    {
        public TestService testService;


        public MainController()
        {
            testService = new TestService();
        }

        public FormValOption formValOption(MainFormOption mainFormOption)
        {
            List<string> tasks = mainFormOption.TodoTask.Items.Cast<string>().ToList();

            FormValOption formValOption = new FormValOption(
                mainFormOption.OutPutFolder.Text.ToString(),
                mainFormOption.ExtractImageCheckBox.Checked,
                mainFormOption.ExtractHyperLinkCheckBox.Checked,
                mainFormOption.ExtractTable.Checked,
                mainFormOption.NotSetMargin.Checked,
                mainFormOption.NotSetPageSize.Checked,
                mainFormOption.TopMargin.Value.ToString(),
                mainFormOption.BottomMargin.Value.ToString(),
                mainFormOption.LeftMargin.Value.ToString(),
                mainFormOption.RightMargin.Value.ToString(),
                mainFormOption.PageWidth.Value.ToString(),
                mainFormOption.PageHeight.Value.ToString(),
                mainFormOption.PageSetOrientation.Text.ToString(),
                mainFormOption.ClearHeader.Checked,
                mainFormOption.ClearFooter.Checked,
                mainFormOption.FirstHeaderFooter.Checked,
                mainFormOption.OddEvenHeaderFooter.Checked,
                mainFormOption.NotSetHeader.Checked,
                mainFormOption.NotSetFooter.Checked,
                mainFormOption.HeaderFontDialog.Font,
                mainFormOption.HeaderAlignComBox.Text.ToString(),
                mainFormOption.HeaderColorDialog.Color,
                mainFormOption.PageHeader.Text.ToString(),
                mainFormOption.FirstHeader.Text.ToString(),
                mainFormOption.OddHeader.Text.ToString(),
                mainFormOption.EvenHeader.Text.ToString(),
                mainFormOption.HeaderImagePath.Text.ToString(),
                mainFormOption.HeaderLine.Checked,
                mainFormOption.FooterFontDialog.Font,
                mainFormOption.FooterAlignComBox.Text.ToString(),
                mainFormOption.FooterColorDialog.Color,
                mainFormOption.PageFooter.Text.ToString(),
                mainFormOption.FirstFooter.Text.ToString(),
                mainFormOption.OddFooter.Text.ToString(),
                mainFormOption.EvenFooter.Text.ToString(),
                mainFormOption.FooterImagePath.Text.ToString(),
                mainFormOption.FooterLine.Checked,
                mainFormOption.PageNumberComBox.Text.ToString(),
                mainFormOption.DocTitle.Text.ToString(),
                mainFormOption.DocSubject.Text.ToString(),
                mainFormOption.DocCategory.Text.ToString(),
                mainFormOption.DocDescription.Text.ToString(),
                mainFormOption.DocCreator.Text.ToString(),
                mainFormOption.DocVersion.Text.ToString(),
                mainFormOption.DocEditPrctCheckBox.Checked,
                mainFormOption.DocEditPrctRemove.Checked,
                mainFormOption.DocEditPassword.Text.ToString(),
                getTextList(mainFormOption.FileGrid),
                tasks,
                getTextList(mainFormOption.ReplaceTextGridView),
                getTextList(mainFormOption.ReplaceLinkGridView),
                mainFormOption.CreateTimeCheckBox.Checked,
                mainFormOption.DocCreateTime.Value,
                mainFormOption.UpdateTimeCheckBox.Checked,
                mainFormOption.DocUpdateTime.Value
            );
            return formValOption;
        }

        public static Dictionary<string, string> getTextList(DataGridView ReplaceTextGridView)
        {
            Dictionary<string, string> lists = new Dictionary<string, string>();

            foreach (DataGridViewRow row in ReplaceTextGridView.Rows)
            {
                string source = row.Cells[0].Value == null ? "" : row.Cells[0].Value.ToString();
                string target = row.Cells[1].Value == null ? "" : row.Cells[1].Value.ToString();
                if (source.Length == 0 && target.Length == 0)
                {
                    continue;
                }
                lists.Add(source, target);
            }
            return lists;
        }

        public void InputFolderBtnEvent(TextBox inputFolder, DataGridView fileGrid)
        {
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                String path = dilog.SelectedPath;
                inputFolder.Text = path;
                DirectoryInfo directoryInfo = new DirectoryInfo(path);
                FileInfo[] files = directoryInfo.GetFiles("*.docx", SearchOption.AllDirectories);

                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                dt.Columns.Add("filename", typeof(string));
                dt.Columns.Add("filepath", typeof(string));
                dt.Columns.Add("filesize", typeof(string));
                dt.Columns.Add("result", typeof(string));
                foreach (FileInfo f in files)
                {
                    string filename = f.Name.Substring(0, f.Name.LastIndexOf("."));
                    string filepath = f.FullName;
                    string filesize = System.Math.Ceiling(f.Length / 1024.0) + " KB";
                    if ((f.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
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

        public void ExportFailFileBtnEvent(DataGridView fileGrid)
        {
            if (fileGrid.DataSource == null || fileGrid.Rows.Count == 0)
            {
                return;
            }
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                DataTable dt = ((DataTable)fileGrid.DataSource).Copy();
                DataView dv = dt.DefaultView;
                string filterString = string.Format("result in ('{0}','{1}','{2}')", ConstData.FAIL, ConstData.protectDocuemnt, ConstData.blankDocuemnt);
                dv.RowFilter = filterString;
                DataTable dataTable = dv.ToTable();
                foreach (DataRow row in dataTable.Rows)
                {
                    string pLocalFilePath = row["filepath"].ToString();
                    string pSaveFilePath = dilog.SelectedPath + @"\" + Path.GetFileName(pLocalFilePath);
                    File.Copy(pLocalFilePath, pSaveFilePath, true);
                }
                System.Diagnostics.Process.Start("explorer.exe", dilog.SelectedPath);
            }
        }


        public void OutputFolderBtnEvent(TextBox outPutFolder)
        {
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                outPutFolder.Text = dilog.SelectedPath;
            }
        }

        public void addToTaskCheck(ListBox todoTask, CheckBox checkBox)
        {
            string title = checkBox.Parent.Text;
            if (checkBox.Checked)
            {
                if (title == ConstData.extractTabText)
                {
                    todoTask.Items.Insert(0, title);
                }
                else
                {
                    todoTask.Items.Add(title);
                }
            }
            else
            {
                todoTask.Items.Remove(title);
            }
        }

        public void NotSetMarginEvent(CheckBox notSetMargin, NumericUpDown topMargin,
            NumericUpDown bottomMargin, NumericUpDown leftMargin, NumericUpDown rightMargin)
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

        public void NotSetPageSizeEvent(CheckBox notSetPageSize, NumericUpDown pageWidth, NumericUpDown pageHeight)
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

        public void test()
        {
            testService.test();
        }



    }
}
