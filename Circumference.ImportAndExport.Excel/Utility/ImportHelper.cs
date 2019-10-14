﻿using Circumference.ImportAndExport.Core;
using Circumference.ImportAndExport.Core.Extension;
using Circumference.ImportAndExport.Core.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Circumference.ImportAndExport.Excel.Utility
{
    public class ImportHelper<T> : IDisposable where T : class, new()
    {
        /// <summary>
        /// </summary>
        /// <param name="filePath"></param>
        public ImportHelper(string filePath = null)
        {
            FilePath = filePath;
        }

        /// <summary>
        ///     导入全局设置
        /// </summary>
        protected ExcelImporterAttribute ExcelImporterAttribute { get; set; }

        /// <summary>
        ///     导入文件路径
        /// </summary>
        protected string FilePath { get; set; }

        /// <summary>
        ///     导入结果
        /// </summary>
        protected ImportResult<T> ImportResult { get; set; }

        /// <summary>
        ///     列头定义
        /// </summary>
        protected List<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }

        public void Dispose()
        {
            ExcelImporterAttribute = null;
            FilePath = null;
            ImporterHeaderInfos = null;
            ImportResult = null;
            GC.Collect();
        }

        /// <summary>
        ///     导入模型验证数据
        /// </summary>
        /// <returns></returns>
        public Task<ImportResult<T>> Import(string filePath = null)
        {
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                FilePath = filePath;
            }

            ImportResult = new ImportResult<T>
            {
                RowErrors = new List<DataRowErrorInfo>()
            };
            ExcelImporterAttribute =
                typeof(T).GetAttribute<ExcelImporterAttribute>(true) ?? new ExcelImporterAttribute();

            try
            {
                CheckImportFile(FilePath);

                using (Stream stream = new FileStream(FilePath, FileMode.Open))
                {
                    using (var excelPackage = new ExcelPackage(stream))
                    {
                        #region 检查模板

                        ParseTemplate(excelPackage);
                        if (ImportResult.HasError)
                        {
                            return Task.FromResult(ImportResult);
                        }

                        #endregion

                        ParseData(excelPackage);
                        #region 数据验证
                        for (var i = 0; i < ImportResult.Data.Count; i++)
                        {
                            ICollection<ValidationResult> validationResults;
                            var isValid = ValidatorHelper.TryValidate(ImportResult.Data[i], out validationResults);
                            if (!isValid)
                            {
                                var dataRowError = new DataRowErrorInfo
                                {
                                    RowIndex = ExcelImporterAttribute.HeaderRowIndex + i + 1
                                };
                                foreach (var validationResult in validationResults)
                                {
                                    var key = validationResult.MemberNames.First();
                                    var column = ImporterHeaderInfos.FirstOrDefault(a => a.PropertyName == key);
                                    if (column != null)
                                    {
                                        key = column.ExporterHeader.Name;
                                    }

                                    var value = validationResult.ErrorMessage;
                                    if (dataRowError.FieldErrors.ContainsKey(key))
                                    {
                                        dataRowError.FieldErrors[key] += Environment.NewLine + value;
                                    }
                                    else
                                    {
                                        dataRowError.FieldErrors.Add(key, value);
                                    }
                                }

                                ImportResult.RowErrors.Add(dataRowError);
                            }
                        }
                        //先不检查重复了，以后再说
                        //RepeatDataCheck();

                        #endregion

                        LabelingError(excelPackage);
                    }
                }
            }
            catch (Exception ex)
            {
                ImportResult.Exception = ex;
            }

            return Task.FromResult(ImportResult);
        }

        /// <summary>
        /// 检查重复数据
        /// </summary>
        //private void RepeatDataCheck()
        //{
        //    //获取需要检查重复数据的列
        //    var notAllowRepeatCols = ImporterHeaderInfos.Where(p => p.ExporterHeader.IsAllowRepeat == false).ToList();
        //    if (notAllowRepeatCols.Count == 0)
        //    {
        //        return;
        //    }

        //    var rowIndex = ExcelImporterAttribute.HeaderRowIndex;
        //    //var qDataList = ImportResult.Data.Select(p =>
        //    //{
        //    //    rowIndex++;
        //    //    return new { RowIndex = rowIndex, RowData = p };
        //    //}).ToList().AsQueryable();
        //    var qDataList = ImportResult.Data.Select(p => new { RowIndex = rowIndex, RowData = p });

        //    foreach (var notAllowRepeatCol in notAllowRepeatCols)
        //    {
        //        //查询指定列
        //        var qDataByProp = qDataList
        //            .Select($"new(RowData.{notAllowRepeatCol.PropertyName} as Value, RowIndex)")
        //            .OrderBy("Value").ToDynamicList();

        //        //重复行的行号
        //        var listRepeatRows = new List<int>();
        //        for (var i = 0; i < qDataByProp.Count; i++)
        //        {
        //            //当前行值
        //            var currentValue = qDataByProp[i].Value;
        //            if (i == 0 || string.IsNullOrEmpty(currentValue?.ToString()))
        //            {
        //                continue;
        //            }

        //            //上一行的值
        //            var preValue = qDataByProp[i - 1].Value;
        //            if (currentValue == preValue)
        //            {
        //                listRepeatRows.Add(qDataByProp[i - 1].RowIndex);
        //                listRepeatRows.Add(qDataByProp[i].RowIndex);
        //                //如果不是最后一行，则继续检测
        //                if (i != qDataByProp.Count - 1)
        //                {
        //                    continue;
        //                }
        //            }

        //            if (listRepeatRows.Count == 0)
        //            {
        //                continue;
        //            }

        //            var errorIndexsStr = string.Join("，", listRepeatRows.Distinct());
        //            foreach (var repeatRow in listRepeatRows.Distinct())
        //            {
        //                var dataRowError = ImportResult.RowErrors.FirstOrDefault(p => p.RowIndex == repeatRow);
        //                if (dataRowError == null)
        //                {
        //                    dataRowError = new DataRowErrorInfo
        //                    {
        //                        RowIndex = repeatRow
        //                    };
        //                    ImportResult.RowErrors.Add(dataRowError);
        //                }

        //                var key = notAllowRepeatCol.ExporterHeader?.Name ??
        //                          notAllowRepeatCol.PropertyName;
        //                var error = $"存在数据重复，请检查！所在行：{errorIndexsStr}。";
        //                if (dataRowError.FieldErrors.ContainsKey(key))
        //                {
        //                    dataRowError.FieldErrors[key] += Environment.NewLine + error;
        //                }
        //                else
        //                {
        //                    dataRowError.FieldErrors.Add(key, error);
        //                }
        //            }

        //            listRepeatRows.Clear();
        //        }
        //    }
        //}

        /// <summary>
        ///     标注错误
        /// </summary>
        /// <param name="excelPackage"></param>
        protected virtual void LabelingError(ExcelPackage excelPackage)
        {
            //是否标注错误
            if (ExcelImporterAttribute.IsLabelingError && ImportResult.HasError)
            {
                var worksheet = GetImportSheet(excelPackage);
                //TODO:标注模板错误
                //标注数据错误
                foreach (var item in ImportResult.RowErrors)
                {
                    foreach (var field in item.FieldErrors)
                    {
                        var col = ImporterHeaderInfos.First(p => p.ExporterHeader.Name == field.Key);
                        var cell = worksheet.Cells[item.RowIndex, col.ExporterHeader.ColumnIndex];
                        cell.Style.Font.Color.SetColor(Color.Red);
                        cell.Style.Font.Bold = true;
                        cell.AddComment(string.Join(",", field.Value), col.ExporterHeader.Author);
                    }
                }

                var ext = Path.GetExtension(FilePath);
                excelPackage.SaveAs(new FileInfo(FilePath.Replace(ext, "_" + ext)));
            }
        }

        /// <summary>
        ///     检查导入文件路劲
        /// </summary>
        /// <exception cref="ArgumentException">文件路径不能为空! - filePath</exception>
        private static void CheckImportFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("文件路径不能为空!", nameof(filePath));
            }

            //TODO:在Docker容器中存在文件路径找不到问题，暂时先注释掉
            //if (!File.Exists(filePath))
            //{
            //    throw new ImportException("导入文件不存在!");
            //}
        }

        /// <summary>
        ///     解析模板
        /// </summary>
        /// <returns></returns>
        protected virtual void ParseTemplate(ExcelPackage excelPackage)
        {
            ImportResult.TemplateErrors = new List<TemplateErrorInfo>();
            Dictionary<int, IDictionary<string, int>> enumColumns;
            List<int> boolColumns;
            //获取导入实体列定义
            ParseImporterHeader(out enumColumns, out boolColumns);
            try
            {
                //根据名称获取Sheet，如果不存在则取第一个
                var worksheet = GetImportSheet(excelPackage);
                var excelHeaders = new Dictionary<string, int>();
                var endColumnCount = ExcelImporterAttribute.EndColumnCount ?? worksheet.Dimension.End.Column;
                for (var columnIndex = 1; columnIndex <= endColumnCount; columnIndex++)
                {
                    var header = worksheet.Cells[ExcelImporterAttribute.HeaderRowIndex, columnIndex].Text;

                    //如果未设置读取的截止列，则默认指定为出现空格，则读取截止
                    if (ExcelImporterAttribute.EndColumnCount.HasValue &&
                        columnIndex > ExcelImporterAttribute.EndColumnCount.Value ||
                        string.IsNullOrWhiteSpace(header))
                    {
                        break;
                    }

                    //不处理空表头
                    if (string.IsNullOrWhiteSpace(header))
                    {
                        continue;
                    }

                    if (excelHeaders.ContainsKey(header))
                    {
                        ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                        {
                            ErrorLevel = ErrorLevels.Error,
                            ColumnName = header,
                            RequireColumnName = null,
                            Message = "列头重复！"
                        });
                    }

                    excelHeaders.Add(header, columnIndex);
                }

                foreach (var item in ImporterHeaderInfos)
                {
                    if (!excelHeaders.ContainsKey(item.ExporterHeader.Name))
                    {
                        //仅验证必填字段
                        if (item.IsRequired)
                        {
                            ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                            {
                                ErrorLevel = ErrorLevels.Error,
                                ColumnName = null,
                                RequireColumnName = item.ExporterHeader.Name,
                                Message = "当前导入模板中未找到此字段！"
                            });
                            continue;
                        }

                        ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                        {
                            ErrorLevel = ErrorLevels.Warning,
                            ColumnName = null,
                            RequireColumnName = item.ExporterHeader.Name,
                            Message = "当前导入模板中未找到此字段！"
                        });
                    }
                    else
                    {
                        item.IsExist = true;
                        //设置列索引
                        if (item.ExporterHeader.ColumnIndex == 0)
                        {
                            item.ExporterHeader.ColumnIndex = excelHeaders[item.ExporterHeader.Name];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                {
                    ErrorLevel = ErrorLevels.Error,
                    ColumnName = null,
                    RequireColumnName = null,
                    Message = $"模板出现未知错误：{ex}"
                });
                throw new Exception($"模板出现未知错误：{ex.Message}", ex);
            }
        }

        /// <summary>
        ///     解析头部
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentException">导入实体没有定义ImporterHeader属性</exception>
        protected virtual bool ParseImporterHeader(out Dictionary<int, IDictionary<string, int>> enumColumns,
            out List<int> boolColumns)
        {
            ImporterHeaderInfos = new List<ImporterHeaderInfo>();
            enumColumns = new Dictionary<int, IDictionary<string, int>>();
            boolColumns = new List<int>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties.Length == 0)
            {
                return false;
            }

            for (var i = 0; i < objProperties.Length; i++)
            {
                //TODO:简化并重构
                //如果不设置，则自动使用默认定义
                var importerHeaderAttribute =
                    (objProperties[i].GetCustomAttributes(typeof(ImporterHeaderAttribute), true) as
                        ImporterHeaderAttribute[])?.FirstOrDefault() ?? new ImporterHeaderAttribute
                        {
                            Name = objProperties[i].GetDisplayName() ?? objProperties[i].Name
                        };

                if (string.IsNullOrWhiteSpace(importerHeaderAttribute.Name))
                {
                    importerHeaderAttribute.Name = objProperties[i].GetDisplayName() ?? objProperties[i].Name;
                }

                ImporterHeaderInfos.Add(new ImporterHeaderInfo
                {
                    IsRequired = objProperties[i].IsRequired(),
                    PropertyName = objProperties[i].Name,
                    ExporterHeader = importerHeaderAttribute
                });
                if (objProperties[i].PropertyType.BaseType?.Name.ToLower() == "enum")
                {
                    enumColumns.Add(i + 1, objProperties[i].PropertyType.GetEnumDisplayNames());
                }

                if (objProperties[i].PropertyType == typeof(bool))
                {
                    boolColumns.Add(i + 1);
                }
            }

            return true;
        }

        /// <summary>
        ///     构建Excel模板
        /// </summary>
        protected virtual void StructureExcel(ExcelPackage excelPackage)
        {
            var worksheet =
                excelPackage.Workbook.Worksheets.Add(typeof(T).GetDisplayName() ??
                                                     ExcelImporterAttribute.SheetName ?? "导入数据");
            Dictionary<int, IDictionary<string, int>> enumColumns;
            List<int> boolColumns;
            if (!ParseImporterHeader(out enumColumns, out boolColumns))
            {
                return;
            }

            //设置列头
            for (var i = 0; i < ImporterHeaderInfos.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = ImporterHeaderInfos[i].ExporterHeader.Name;
                if (!string.IsNullOrWhiteSpace(ImporterHeaderInfos[i].ExporterHeader.Description))
                {
                    worksheet.Cells[1, i + 1].AddComment(ImporterHeaderInfos[i].ExporterHeader.Description,
                        ImporterHeaderInfos[i].ExporterHeader.Author);
                }
                //如果必填，则列头标红
                if (ImporterHeaderInfos[i].IsRequired)
                {
                    worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.Red);
                }
            }

            worksheet.Cells.AutoFitColumns();
            worksheet.Cells.Style.WrapText = true;
            worksheet.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[worksheet.Dimension.Address].Style.Fill.BackgroundColor.SetColor(Color.DarkSeaGreen);

            //枚举处理
            foreach (var enumColumn in enumColumns)
            {
                var range = ExcelCellBase.GetAddress(1, enumColumn.Key, ExcelPackage.MaxRows, enumColumn.Key);
                var dataValidations = worksheet.DataValidations.AddListValidation(range);
                foreach (var displayName in enumColumn.Value)
                {
                    dataValidations.Formula.Values.Add(displayName.Key);
                }
            }

            //Bool类型处理
            foreach (var boolColumn in boolColumns)
            {
                var range = ExcelCellBase.GetAddress(1, boolColumn, ExcelPackage.MaxRows, boolColumn);
                var dataValidations = worksheet.DataValidations.AddListValidation(range);
                dataValidations.Formula.Values.Add("是");
                dataValidations.Formula.Values.Add("否");
            }
        }

        /// <summary>
        ///     解析数据
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentException">最大允许导入条数不能超过5000条</exception>
        protected virtual void ParseData(ExcelPackage excelPackage)
        {
            var worksheet = GetImportSheet(excelPackage);
            if (worksheet.Dimension.End.Row > 5000)
            {
                throw new ArgumentException("最大允许导入条数不能超过5000条");
            }

            ImportResult.Data = new List<T>();
            var propertyInfos = new List<PropertyInfo>(typeof(T).GetProperties());

            for (var rowIndex = ExcelImporterAttribute.HeaderRowIndex + 1;
                rowIndex <= worksheet.Dimension.End.Row;
                rowIndex++)
            {
                var isNullNumber = 1;
                for (var column = 1; column < worksheet.Dimension.End.Column; column++)
                {
                    if (worksheet.Cells[rowIndex, column].Text == string.Empty)
                    {
                        isNullNumber++;
                    }
                }

                if (isNullNumber < worksheet.Dimension.End.Column)
                {
                    var dataItem = new T();
                    foreach (var propertyInfo in propertyInfos)
                    {
                        var col = ImporterHeaderInfos.Find(a => a.PropertyName == propertyInfo.Name);
                        //检查Excel中是否存在
                        if (!col.IsExist)
                        {
                            continue;
                        }

                        var cell = worksheet.Cells[rowIndex, col.ExporterHeader.ColumnIndex];
                        try
                        {
                            var cellValue = cell.Value?.ToString();
                            if (cellValue == null) throw new ArgumentException();
                            switch (propertyInfo.PropertyType.BaseType?.Name)
                            {
                                case "Enum":
                                    var enumDisplayNames = propertyInfo.PropertyType.GetEnumDisplayNames();
                                    if (enumDisplayNames.ContainsKey(cellValue))
                                    {
                                        propertyInfo.SetValue(dataItem,
                                            enumDisplayNames[cellValue]);
                                    }
                                    else
                                    {
                                        AddRowDataError(rowIndex, col, $"值 {cell.Value} 不存在模板下拉选项中");
                                    }

                                    continue;

                            }

                            //var cellValue = cell.Value?.ToString();
                            switch (propertyInfo.PropertyType.GetCSharpTypeName())
                            {
                                case "Boolean":
                                    propertyInfo.SetValue(dataItem, GetBooleanValue(cellValue));
                                    break;
                                case "Nullable<Boolean>":
                                    propertyInfo.SetValue(dataItem,
                                        string.IsNullOrWhiteSpace(cellValue)
                                             ? (bool?)null
                                            : GetBooleanValue(cellValue));
                                    break;
                                case "String":
                                    //TODO:进一步优化
                                    //移除所有的空格，包括中间的空格
                                    if (col.ExporterHeader.FixAllSpace)
                                    {
                                        propertyInfo.SetValue(dataItem, cellValue?.Replace(" ", string.Empty));
                                    }
                                    else if (col.ExporterHeader.AutoTrim)
                                    {
                                        propertyInfo.SetValue(dataItem, cellValue?.Trim());
                                    }
                                    else
                                    {
                                        propertyInfo.SetValue(dataItem, cellValue);
                                    }

                                    break;
                                //long
                                case "Int64":
                                    {
                                        Int64 number;
                                        if (!long.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int64>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        Int64 number;
                                        if (!long.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Int32":
                                    {
                                        Int32 number;
                                        if (!int.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int32>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        Int32 number;
                                        if (!int.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Int16":
                                    {
                                        Int16 number;
                                        if (!short.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int16>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        Int16 number;
                                        if (!short.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Decimal":
                                    {
                                        decimal number;
                                        if (!decimal.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Decimal>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        decimal number;
                                        if (!decimal.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Double":
                                    {
                                        double number;
                                        if (!double.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Double>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        double number;
                                        if (!double.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                //case "float":
                                case "Single":
                                    {
                                        float number;
                                        if (!float.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Single>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        float number;
                                        if (!float.TryParse(cellValue, out number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "DateTime":
                                    {
                                        DateTime date;
                                        if (!DateTime.TryParse(cellValue, out date))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的日期时间格式！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, date);
                                    }
                                    break;
                                case "Nullable<DateTime>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        DateTime date;
                                        if (!DateTime.TryParse(cellValue, out date))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的日期时间格式！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, date);
                                    }
                                    break;
                                default:
                                    propertyInfo.SetValue(dataItem, cell.Value);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            AddRowDataError(rowIndex, col, ex.Message);
                        }
                    }

                    ImportResult.Data.Add(dataItem);
                }
            }
        }

        /// <summary>
        ///     获取导入的Sheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        protected virtual ExcelWorksheet GetImportSheet(ExcelPackage excelPackage) => excelPackage.Workbook.Worksheets[typeof(T).GetDisplayName()] ??
                   excelPackage.Workbook.Worksheets[ExcelImporterAttribute.SheetName] ??
                   excelPackage.Workbook.Worksheets[1];

        /// <summary>
        ///     添加数据行错误
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="importerHeaderInfo"></param>
        /// <param name="errorMessage"></param>
        protected virtual void AddRowDataError(int rowIndex, ImporterHeaderInfo importerHeaderInfo,
            string errorMessage = "数据格式无效！")
        {
            var rowError = ImportResult.RowErrors.FirstOrDefault(p => p.RowIndex == rowIndex);
            if (rowError == null)
            {
                rowError = new DataRowErrorInfo
                {
                    RowIndex = rowIndex
                };
                ImportResult.RowErrors.Add(rowError);
            }

            rowError.FieldErrors.Add(importerHeaderInfo.ExporterHeader.Name, errorMessage);
        }

        /// <summary>
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static bool GetBooleanValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            switch (value.ToLower())
            {
                case "1":
                case "是":
                case "yes":
                case "true":
                    return true;
                case "0":
                case "否":
                case "no":
                case "false":
                default:
                    return false;
            }
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        public Task<byte[]> GenerateTemplateByte()
        {
            ExcelImporterAttribute =
                typeof(T).GetAttribute<ExcelImporterAttribute>(true) ?? new ExcelImporterAttribute();
            using (var excelPackage = new ExcelPackage())
            {
                StructureExcel(excelPackage);
                return Task.FromResult(excelPackage.GetAsByteArray());
            }
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">文件名必须填写! - fileName</exception>
        public Task<TemplateFileInfo> GenerateTemplate(string fileName = null)
        {
            ExcelImporterAttribute =
                typeof(T).GetAttribute<ExcelImporterAttribute>(true) ?? new ExcelImporterAttribute();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("文件名必须填写!", fileName);
            }

            var fileInfo =
                ExcelHelper.CreateExcelPackage(fileName, excelPackage => StructureExcel(excelPackage));
            return Task.FromResult(fileInfo);
        }
    }
}
