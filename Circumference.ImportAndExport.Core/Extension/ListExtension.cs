using Circumference.ImportAndExport.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Core.Extension
{
    public static class ListExtension
    {
        /// <summary>
        /// 将List集合转成DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> list)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p =>
                new DataColumn(p.PropertyType.GetAttribute<ExporterAttribute>()?.Name ?? p.GetDisplayName() ?? p.Name,
                    p.PropertyType)).ToArray());
            if (list.Count <= 0) return dt;

            for (var i = 0; i < list.Count; i++)
            {
                var tempList = new ArrayList();
                foreach (var obj in props.Select(pi => pi.GetValue(list.ElementAt(i), null))) tempList.Add(obj);
                var array = tempList.ToArray();
                dt.LoadDataRow(array, true);
            }

            return dt;
        }
    }
}
