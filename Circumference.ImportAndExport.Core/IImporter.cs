using Circumference.ImportAndExport.Core.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core
{
    /// <summary>
    ///     导入
    /// </summary>
    public interface IImporter
    {
        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<TemplateFileInfo> GenerateTemplate<T>(string fileName) where T : class, new();

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        Task<byte[]> GenerateTemplateByte<T>() where T : class, new();

        /// <summary>
        ///     导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(string filePath) where T : class, new();
    }
}
