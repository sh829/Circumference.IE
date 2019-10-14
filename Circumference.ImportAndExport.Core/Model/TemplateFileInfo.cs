using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core.Model
{
    public class TemplateFileInfo
    {
        public TemplateFileInfo()
        {
        }

        public TemplateFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }

        public string FileName { get; set; }

        public string FileType { get; set; }
    }
}
