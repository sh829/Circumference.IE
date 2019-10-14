﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExporterAttribute : Attribute
    {
        /// <summary>
        ///     名称(比如当前Sheet 名称)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///     头部字体大小
        /// </summary>
        public float? HeaderFontSize { set; get; }

        /// <summary>
        ///     正文字体大小
        /// </summary>
        public float? FontSize { set; get; }
    }
}
