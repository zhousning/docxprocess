using System;
using System.Collections.Generic;
using System.ComponentModel;
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

        public void Process(BackgroundWorker bw, List<string> tasks, ref DataGridView fileGrid, TextBox outPutFolder, ref Button startProcess, TaskProcess taskProcess)
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

            foreach (DataGridViewRow row in fileGrid.Rows)
            {
                if (!bw.CancellationPending)
                {
                    string filepath = row.Cells["filepath"].Value.ToString();
                    string filename = row.Cells["filename"].Value.ToString();
                    string result = "";
                    try
                    {
                        if (!String.IsNullOrWhiteSpace(filepath) && File.Exists(filepath))
                        {
                            string targetFile = outputDirectory + @"\" + filename;
                            taskProcess(tasks, filepath, targetFile, ref result);
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
                        row.Cells["result"].Value = result;
                    }
                }
            }
        }
    }
}
