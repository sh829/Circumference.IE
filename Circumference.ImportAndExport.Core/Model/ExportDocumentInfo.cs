using Circumference.ImportAndExport.Core.Extension;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core.Model
{
    public class ExportDocumentInfo<TData> where TData : class
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfo(IList<TData> datas)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Datas = datas;
            Title = typeof(TData).GetAttribute<ExporterAttribute>()?.Name ?? typeof(TData).Name;

            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exporterHeader = propertyInfo.PropertyType.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public IList<TData> Datas { get; set; }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        public DataTable ToDataTable()
        {
            return Datas.ToDataTable();
        }
    }
}
