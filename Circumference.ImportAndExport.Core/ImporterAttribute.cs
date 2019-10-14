using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ImporterAttribute : Attribute
    {
        /// <summary>
        ///     表头位置
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;
    }
}
