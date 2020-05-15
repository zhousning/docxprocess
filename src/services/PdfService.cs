using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.Word;

namespace Docx.src.services
{
    class PdfService
    {
       public bool WordToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            /*Application application = new Application();
            application.Visible = false;
            Document document = null;
            try
            {
                document = application.Documents.Open(sourcePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                object wdFormatPDFDocument = WdSaveFormat.wdFormatPDF;
                document.SaveAs2(targetPath, wdFormatPDFDocument, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);               
                result = true;
            }
            catch (COMException exception)
            {
                result = false;
            }
            finally
            {
                document.Close(Missing.Value, Missing.Value, Missing.Value);
                application.Quit(Missing.Value, Missing.Value, Missing.Value);
            }*/
            return result;
        }

        /*public bool WordToPdfWithWPS(string sourcePath, string targetPath)
        {
            WPS.ApplicationClass app = new WPS.ApplicationClass();
            WPS.Document doc = null;
            try
            {
                doc = app.Documents.Open(sourcePath, true, true, false, null, null, false, "", null, 100, 0, true, true, 0, true);
                doc.ExportPdf(targetPath, "", "");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
                doc.Close();
            }
            return true;
        }*/
    }
    }
