using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Docx.src.services
{
    class BaseService
    {
        public bool IsMatch(string expression, string str)
        {
            Regex reg = new Regex(expression);
            if (string.IsNullOrWhiteSpace(str))
                return false;
            return reg.IsMatch(str);
        }

        public string makedir(string path)
        {
            string source = @path + @"资源\";
            bool flag = Directory.Exists(source);
            if (!flag)
            {
                Directory.CreateDirectory(source);
            }
            return source;
        }
    }
}
