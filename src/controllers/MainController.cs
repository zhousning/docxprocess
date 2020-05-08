using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Docx.src.model;
using Docx.src.services;
using Xceed.Words.NET;

namespace Docx.src.controllers
{
    class MainController
    {

        public MainController()
        {
            
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
            if (checkBox.Checked)
            {
                todoTask.Items.Add(checkBox.Parent.Text);
            }
            else
            {
                todoTask.Items.Remove(checkBox.Parent.Text);
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


        
    }
}
