using System;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Docx.src.docxprocess;
using Xceed.Words.NET;
using System.Drawing;
using Docx.src.model;
using Xceed.Document.NET;
using System.IO;
using System.Drawing.Imaging;

namespace Docx.src.services
{
    class PageSettingService
    {
        private PageSetting pageSetting;

        public PageSettingService()
        {
            this.pageSetting = new PageSetting();
        }
        public void marginSetting(DocX document, Boolean notSetMargin, string topMargin, string bottomMargin, string leftMargin, string rightMargin)
        {
            if (!notSetMargin)
            {
                float marginTop = float.Parse(topMargin) * ConstData.POUND;
                float marginBottom = float.Parse(bottomMargin) * ConstData.POUND;
                float marginLeft = float.Parse(leftMargin) * ConstData.POUND;
                float marginRight = float.Parse(rightMargin) * ConstData.POUND;
                this.pageSetting.Margins(document, marginTop, marginBottom, marginLeft, marginRight);
            }
        }

        public void pageSizeSetting(DocX document, Boolean notSetPageSize, string pageWidth, string pageHeight)
        {
            if (!notSetPageSize)
            {
                float width = float.Parse(pageWidth) * ConstData.POUND;
                float height = float.Parse(pageHeight) * ConstData.POUND;
                this.pageSetting.pageSize(document, width, height);
            }
        }

        public void pageOrientation(DocX document, string orientation)
        {
            this.pageSetting.pageOrientation(document, orientation);
        }

        public void test()
        {
            using (var document = DocX.Load(@"C:\Users\周宁\Desktop\新建文件夹 (3)\最终版 28版.docx"))
            {
                // Add a title
                document.InsertParagraph(0, "Modifying Image by adding text/circle into the following image", false).FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

                var images = document.Images;
                for(int i=0; i<images.Count; i++)
                {
                    Bitmap bitmap;

                    using (var stream = images[i].GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        bitmap = new Bitmap(stream);
                        bitmap.Save(@"C:\Users\周宁\Desktop\新建文件夹 (3)\" + i + ".png", ImageFormat.Png);
                    }
                }
                // Get the first image in the document.
                /*var image = document.Images.FirstOrDefault();
                if (image != null)
                {
                    // Create a bitmap from the image.
                    Bitmap bitmap;
                    using (var stream = image.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        bitmap = new Bitmap(stream);
                    }
                    // Get the graphic from the bitmap to be able to draw in it.
                    var graphic = Graphics.FromImage(bitmap);
                    if (graphic != null)
                    {
                        // Draw a string with a specific font, font size and color at (0,10) from top left of the image.
                        graphic.DrawString("@copyright", new System.Drawing.Font("Arial Bold", 12), Brushes.Red, new PointF(0f, 10f));
                        // Draw a blue circle of 10x10 at (30, 5) from the top left of the image.
                        graphic.FillEllipse(Brushes.Blue, 30, 5, 10, 10);

                        // Save this Bitmap back into the document using a Create\Write stream.
                        bitmap.Save(image.GetStream(FileMode.Create, FileAccess.Write), ImageFormat.Png);
                    }
                }*/
                //document.SaveAs(@"C:\Users\周宁\Desktop\新建文件夹 (2)\最终版 29版.docx");

                Console.WriteLine("\tCreated: ModifyImage.docx\n");
            }
        }

       
    }
}
