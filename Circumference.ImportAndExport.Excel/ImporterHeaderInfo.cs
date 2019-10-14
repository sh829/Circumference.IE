using Circumference.ImportAndExport.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Excel
{
    /// <summary>
    ///     导入列头设置
    /// </summary>
    public class ImporterHeaderInfo
    {
        /// <summary>
        ///     是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        ///     列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        ///     列属性
        /// </summary>
        public ImporterHeaderAttribute ExporterHeader { get; set; }

        /// <summary>
        ///     是否存在
        /// </summary>
        internal bool IsExist { get; set; }
    }
}
