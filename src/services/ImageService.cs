using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace Docx.src.services
{
    class ImageService:BaseService
    {
        public void ImageSet(DocX document, string output, CheckBox ExtractImageCheckBox)
        {
            if (ExtractImageCheckBox.Checked)
            {
                extractImages(document, output);
            }
        }
        public void extractImages(DocX document, string output)
        {
            var images = document.Images;
            string outputPath = makedir(output);
            for (int i = 0; i < images.Count; i++)
            {
                Bitmap bitmap;

                using (var stream = images[i].GetStream(FileMode.Open, FileAccess.ReadWrite))
                {
                    bitmap = new Bitmap(stream);
                    bitmap.Save(outputPath + "$" + i + ".png", ImageFormat.Png);
                }
            }

            var paragraphs = document.Paragraphs;
            var foldername = Path.GetDirectoryName(outputPath);
            var txtname = Path.GetFileNameWithoutExtension(foldername);
            var picture_start = 0;
            for (int j = 0; j < paragraphs.Count; j++)
            {
                var p = paragraphs[j];
                StreamWriter dout = new StreamWriter(outputPath  + txtname + ".txt", true);
                var pictures = p.Pictures;
                var picture_str = "";
                var number = 0;
                for (int z =0; z < pictures.Count; z++)
                {
                    number = z + 1;
                    var img_no = z + picture_start;
                    picture_str += "$" + img_no + System.Environment.NewLine;
                }
                picture_start += number;
                dout.Write(picture_str + p.Text);
                dout.Write(System.Environment.NewLine); //换行
                dout.Close();
            }
        }

        public string ImgToBase64String(string Imagefilename)
        {
            try
            {
                Bitmap bmp = new Bitmap(Imagefilename);

                MemoryStream ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
                return Convert.ToBase64String(arr);
            }
            catch (Exception ex)
            {
                return null;
            }
        }


    }
}
