using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Docx.src.model;
using Docx.src.services;
using NLog;
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
        private Logger logger;

        public delegate void TaskProcess(FormValOption formValOption, string filepath, string targetFile, ref string result);

        public BackWorker(MainFormOption mainFormOption)
        {
            this.pageSettingService = new PageSettingService();
            this.headerFooterService = new HeaderFooterService();
            this.docInfoService = new DocInfoService();
            this.textReplaceService = new TextReplaceService();
            this.paragraphService = new ParagraphService();
            this.imageService = new ImageService();
            this.hyperLinkService = new HyperLinkService();
            this.tableService = new TableService();
            this.pdfService = new PdfService();
            this.mainFormOption = mainFormOption;
            this.logger = LogManager.GetCurrentClassLogger();
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
            FormValOption formValOption = (FormValOption)e.Argument;
            string title = formValOption.ProcessTitle;

            if (title == ConstData.START_PRC)
            {
                e.Result = TaskProcessAsync(bw, formValOption);
            }
            else if (title == ConstData.PDF_EXPORT)
            {
                e.Result = PdfExportAsync(bw, formValOption);
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
            Thread.Sleep(1000);
            this.mainFormOption.ToolStripProgressBar.Value = 0;
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string[] arr = e.UserState.ToString().Split(',');
            int i = int.Parse(arr[0]);
            string result = arr[1];
            this.mainFormOption.FileGrid.Rows[i].Cells["result"].Value = result;
            this.mainFormOption.ToolStripProgressBar.Value = e.ProgressPercentage;
        }

        private int TaskProcessAsync(BackgroundWorker bw, FormValOption formValOption)
        {
            if (formValOption.TodoTask.Count == 0)
            {
                MessageBox.Show("当前没有待处理任务", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return ConstData.STOP_PROCESS;
            }

            return Process(bw, formValOption, ToDoTaskProcess);
        }
        private int PdfExportAsync(BackgroundWorker bw, FormValOption formValOption)
        {
            return Process(bw, formValOption, PDFConverterProcess);
        }



        public int Process(BackgroundWorker bw, FormValOption formValOption, TaskProcess taskProcess)
        {
            string outputDirectory = formValOption.OutPutFolder;
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
            Dictionary<string, string> rows = formValOption.FileGrid;
            int count = rows.Count;
            int i = 0;
            foreach (KeyValuePair<string, string> kv in rows)
            {
                if (!bw.CancellationPending)
                {
                    i++;
                    string filename = kv.Key;
                    string filepath = kv.Value;
                    string result = "";
                    try
                    {
                        if (!String.IsNullOrWhiteSpace(filepath) && File.Exists(filepath))
                        {
                            string targetFile = outputDirectory + @"\" + filename;
                            taskProcess(formValOption, filepath, targetFile, ref result);
                            result = ConstData.SUCCESS;

                        }
                        else
                        {
                            result = ConstData.WITHOUT_FILE;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error("***" + filename + ":  " + ex.Message);
                        result = ConstData.FAIL;
                    }
                    finally
                    {
                        string data = (i - 1) + "," + result;
                        bw.ReportProgress(i * 100 / count, data);
                        Thread.Sleep(1000);
                    }
                }
            }
            return ConstData.FINISH_PROCESS;
        }

        private void ToDoTaskProcess(FormValOption formValOption, string filepath, string targetFile, ref string result)
        {
            targetFile += ConstData.DOCXPREF;
            formValOption.OutPutFolder = targetFile;

            using (var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                using (var document = DocX.Load(fs))
                {
                    List<string> tasks = formValOption.TodoTask;
                    foreach (string title in tasks)
                    {
                        if (title == ConstData.pageSettingTabText)
                        {
                            ProcessWorker.PageSet(document, pageSettingService, formValOption);

                        }
                        else if (title == ConstData.headerFooterTabText)
                        {
                            ProcessWorker.HeaderFooterSet(document, headerFooterService, formValOption);

                        }
                        else if (title == ConstData.docInfoTabText)
                        {
                            ProcessWorker.DocInfoSet(document, docInfoService, formValOption);

                        }
                        else if (title == ConstData.textReplaceTabText)
                        {
                            ProcessWorker.TextReplaceSet(document, textReplaceService, formValOption);

                        }
                        else if (title == ConstData.paragraphTabText)
                        {
                            //ProcessWorker.ParagraphSet(document,);
                        }
                        else if (title == ConstData.extractTabText)
                        {
                            ProcessWorker.ExtractSet(document, imageService, hyperLinkService, tableService, formValOption);

                        }
                    }

                    if (!(tasks.Contains(ConstData.extractTabText) && tasks.Count == 1))
                    {
                        document.SaveAs(targetFile);
                    }
                    if (tasks.Contains(ConstData.docInfoTabText))
                    {
                        ProcessWorker.UpdateFileTime(docInfoService, targetFile, formValOption);

                    }
                    result = ConstData.SUCCESS;
                }
            }
        }

        private void PDFConverterProcess(FormValOption formValOption, string filepath, string targetFile, ref string result)
        {
            targetFile += ConstData.PDFPREF;

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


        private void stopChangeBtn(MainFormOption mainFormOption)
        {
            mainFormOption.TaskProcessBtn.Enabled = true;
            mainFormOption.PdfExportBtn.Enabled = true;
            mainFormOption.OutputFolderBtn.Enabled = true;
            mainFormOption.InputFolderBtn.Enabled = true;
            mainFormOption.ExportFailFile.Enabled = true;
            mainFormOption.StopWork.Enabled = false;
            mainFormOption.FileGrid.AllowUserToDeleteRows = true;
        }


    }
}
