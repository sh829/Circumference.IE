using Circumference.ImportAndExport.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Excel
{
    public class ExporterHeaderInfo
    {
        /// <summary>
        ///     列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///     列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        ///     列属性
        /// </summary>
        public ExporterHeaderAttribute ExporterHeader { get; set; }
    }
}
