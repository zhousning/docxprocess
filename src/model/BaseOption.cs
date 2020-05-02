using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace Docx.src.model
{
    class BaseOption
    {
        public Alignment setAlignment(string alignmentString)
        {
            Alignment alignment= Alignment.left;
            switch (alignmentString)
            {
                case ConstData.ALIGNLEFT:
                    alignment = Alignment.left;
                    break;
                case ConstData.ALIGNCENTER:
                    alignment = Alignment.center;
                    break;
                case ConstData.ALIGNRIGHT:
                    alignment = Alignment.right;
                    break;
                case ConstData.ALIGNBOTH:
                    alignment = Alignment.both;
                    break;
            }
            return alignment;
        }
    }
}
