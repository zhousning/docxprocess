using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Docx.src.services
{
    class MainFormService
    {
        public delegate void TaskProcess(List<string> tasks, string filepath, string targetFile, ref string result);

        public void Process(List<string> tasks, ref DataGridView fileGrid, TextBox outPutFolder,ref Button startProcess, TaskProcess taskProcess)
        {
            string outputDirectory = outPutFolder.Text;
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                MessageBox.Show("请选择输出目录", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DialogResult dialogResult = MessageBox.Show("开始处理前将关闭所有word文件", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (dialogResult == DialogResult.Yes)
            {
                foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
                {
                    p.Kill();
                }
            }
            else
            {
                return;
            }
            string btnTitle = startProcess.Text;
            startProcess.Text = "处理中...";
            startProcess.Enabled = false;
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
                    if (!String.IsNullOrWhiteSpace(filepath) && File.Exists(filepath))
                    {
                        string targetFile = outputDirectory + @"\" + filename;
                        taskProcess(tasks, filepath, targetFile,ref result);
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        result = ConstData.WITHOUT_FILE;
                    }
                }
                catch (Exception ex)
                {
                    result = ConstData.FAIL;
                }
                finally
                {
                    /*DataRow dataRow = dt.NewRow();
                    dataRow["filename"] = filename;
                    dataRow["filepath"] = filepath;
                    dataRow["filesize"] = filesize;
                    dataRow["result"] = result;
                    dt.Rows.Add(dataRow);*/
                    row.Cells["result"].Value = result;
                }
            }
            //ds.Tables.Add(dt);
            //fileGrid.DataSource = ds.Tables[0];
            startProcess.Text = btnTitle;
            startProcess.Enabled = true;
            //MessageBox.Show("处理完毕", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
    }
}
