using Circumference.ImportAndExport.Core.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core
{
    /// <summary>
    ///     导出
    /// </summary>
    public interface IExporter
    {
        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="dataItems">数据</param>
        /// <returns>文件</returns>
        Task<TemplateFileInfo> Export<T>(string fileName, IList<T> dataItems) where T : class;

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray<T>(IList<T> dataItems) where T : class;

        /// <summary>
        ///     导出Excel表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <param name="globalStyle">全局样式</param>
        /// <param name="styles">样式</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName, ExcelHeadStyle globalStyle = null,
            List<ExcelHeadStyle> styles = null);

        /// <summary>
        ///     导出Excel表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class;
    }
}
