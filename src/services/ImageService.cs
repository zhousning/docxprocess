using System;
using System.Collections.Generic;
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
                    bitmap.Save(outputPath + i + ".png", ImageFormat.Png);
                }
            }
        }

        
    }
}
