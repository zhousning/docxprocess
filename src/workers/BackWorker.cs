using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Docx.src.model;
using Docx.src.services;
using Xceed.Words.NET;

namespace Docx.src.workers
{
    class BackWorker
    {
        private PageSettingService pageSettingService;
        private HeaderFooterService headerFooterService;
        private DocInfoService docInfoService;
        private TextReplaceService textReplaceService;
        private ParagraphService paragraphService;
        private ImageService imageService;
        private HyperLinkService hyperLinkService;
        private TableService tableService;
        private PdfService pdfService;
        

        private MainFormOption mainFormOption;
        public delegate void TaskProcess(List<string> tasks, string filepath, string targetFile, ref string result);

        public BackWorker(MainFormOption mainFormOption)
        {
            this.mainFormOption = mainFormOption;

            this.pageSettingService = new PageSettingService();
            this.headerFooterService = new HeaderFooterService();
            this.docInfoService = new DocInfoService();
            this.textReplaceService = new TextReplaceService();
            this.paragraphService = new ParagraphService();
            this.imageService = new ImageService();
            this.hyperLinkService = new HyperLinkService();
            this.tableService = new TableService();
            this.pdfService = new PdfService();
        }

        public BackgroundWorker getWorker()
        {
            BackgroundWorker bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.DoWork += new DoWorkEventHandler(DoWork);
            bgWorker.ProgressChanged += new ProgressChangedEventHandler(ProgressChanged);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunWorkerCompleted);
            return bgWorker;
        }



        private void DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = sender as BackgroundWorker;
             
            foreach (DataGridViewRow row in this.mainFormOption.FileGrid.Rows)
            {
                row.Cells["result"].Value = "";
            }
            string title = (string)e.Argument;

            if (title == ConstData.START_PRC)
            {
                e.Result = TaskProcessAsync(bw);
            }
            else if (title == ConstData.PDF_EXPORT)
            {
                e.Result = PdfExportAsync(bw);
            }

            if (bw.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        private void RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            stopChangeBtn(this.mainFormOption);
            if (e.Cancelled)
            {
                MessageBox.Show("已停止处理");
            }
            else if (e.Error != null)
            {
                string msg = String.Format("An error occurred: {0}", e.Error.Message);
                MessageBox.Show(msg);
            }
            else
            {
                if ((int)e.Result == ConstData.FINISH_PROCESS)
                {
                    string outpath = this.mainFormOption.OutPutFolder.Text;
                    string v_OpenFolderPath = @outpath;
                    System.Diagnostics.Process.Start("explorer.exe", v_OpenFolderPath);
                }
            }
            Thread.Sleep(2000);
            this.mainFormOption.ToolStripProgressBar.Value = 0;
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.mainFormOption.ToolStripProgressBar.Value = e.ProgressPercentage;
        }

        private int TaskProcessAsync(BackgroundWorker bw)
        {
            List<string> tasks = this.mainFormOption.TodoTask.Items.Cast<string>().ToList();
            if (tasks.Count == 0)
            {
                MessageBox.Show("当前没有待处理任务", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return ConstData.STOP_PROCESS;
            }

            return Process(bw, tasks, mainFormOption, ToDoTaskProcess);
        }
        private int PdfExportAsync(BackgroundWorker bw)
        {
            return Process(bw, null, mainFormOption, PDFConverterProcess);
        }

        private void stopChangeBtn(MainFormOption mainFormOption)
        {
            mainFormOption.TaskProcessBtn.Enabled = true;
            mainFormOption.PdfExportBtn.Enabled = true;
            mainFormOption.OutputFolderBtn.Enabled = true;
            mainFormOption.InputFolderBtn.Enabled = true;
            mainFormOption.StopWork.Enabled = false;
            mainFormOption.FileGrid.AllowUserToDeleteRows = true;
        }


        public int Process(BackgroundWorker bw, List<string> tasks, MainFormOption mainFormOption, TaskProcess taskProcess)
        {
            string outputDirectory = mainFormOption.OutPutFolder.Text;
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                MessageBox.Show("请选择输出目录", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return ConstData.STOP_PROCESS;
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
                return ConstData.STOP_PROCESS;
            }
            int count = mainFormOption.FileGrid.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                DataGridViewRow row = mainFormOption.FileGrid.Rows[i];
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
                            bw.ReportProgress((i + 1) * 100 / count);
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
            return ConstData.FINISH_PROCESS;
        }

        private void ToDoTaskProcess(List<string> tasks, string filepath, string targetFile, ref string result)
        {
            targetFile += ConstData.DOCXPREF;
            using (var document = DocX.Load(filepath))
            {
                foreach (string title in tasks)
                {
                    if (title == ConstData.pageSettingTabText)
                    {
                        ProcessWorker.PageSet(document, pageSettingService, mainFormOption);
                    }
                    else if (title == ConstData.headerFooterTabText)
                    {
                        ProcessWorker.HeaderFooterSet(document,headerFooterService, mainFormOption);
                    }
                    else if (title == ConstData.docInfoTabText)
                    {
                        ProcessWorker.DocInfoSet(document,docInfoService,mainFormOption);
                    }
                    else if (title == ConstData.textReplaceTabText)
                    {
                        ProcessWorker.TextReplaceSet(document,textReplaceService,mainFormOption);
                    }
                    else if (title == ConstData.paragraphTabText)
                    {
                        //ProcessWorker.ParagraphSet(document,);
                    }
                    else if (title == ConstData.extractTabText)
                    {
                        ProcessWorker.ExtractSet(document, imageService, hyperLinkService,tableService, mainFormOption);
                    }
                }

                document.SaveAs(targetFile);
                if (tasks.Contains(ConstData.docInfoTabText))
                {
                    ProcessWorker.UpdateFileTime(docInfoService, targetFile, mainFormOption);
                }
                result = ConstData.SUCCESS;
            }
        }

        private void PDFConverterProcess(List<string> tasks, string filepath, string targetFile, ref string result)
        {
            Boolean flag = this.pdfService.WordToPDF(filepath, targetFile);
            if (flag)
            {
                result = ConstData.SUCCESS;
            }
            else
            {
                result = ConstData.FAIL;
            }
        }

       
    }
}
