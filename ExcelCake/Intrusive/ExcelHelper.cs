using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelCake.Intrusive
{
    public class ExcelHelper
    {
        private static readonly string _DateTimeFormatStr = "yyyyMMddHHmmssfff";

        public static IEnumerable<T> GetList<T>(FileInfo file, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            IEnumerable<T> list = new List<T>();

            if (!file.Exists)
            {
                return list;
            }

            //读取excel
            using (ExcelPackage ep = new ExcelPackage(file))
            {
                List<ExcelWorksheet> sheets = ep.Workbook.Worksheets.ToList();
                if (noImportSheets != null)
                {
                    sheets = sheets.Where(p => !noImportSheets.Contains(p.Name)).ToList();
                }

                if (importSheets != null)
                {
                    sheets = sheets.Where(p => importSheets.Contains(p.Name)).ToList();
                }

                if (!string.IsNullOrEmpty(noImportSheetsRegex))
                {
                    sheets = sheets.Where(p => !Regex.IsMatch(p.Name, noImportSheetsRegex)).ToList();
                }

                if (!string.IsNullOrEmpty(importSheetsRegex))
                {
                    sheets = sheets.Where(p => Regex.IsMatch(p.Name, importSheetsRegex)).ToList();
                }

                list = GetCollectionFromSheets<T>(sheets,ep.File.Name);
            }
            return list;
        }

        public static IEnumerable<T> GetList<T>(string filePath, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            return GetList<T>(new FileInfo(filePath), importSheets, noImportSheets, importSheetsRegex, noImportSheetsRegex);
        }

        private static IEnumerable<T> GetCollectionFromSheets<T>(List<ExcelWorksheet> sheets,string excelName) where T : ExcelBase, new()
        {
            var list = new List<T>();

            if (sheets == null || sheets.Count == 0)
            {
                return list;
            }

            Type type = typeof(T);

            var importSetting = new ImportExcelSetting(type);
            if (importSetting.ImportColumns.Count > 0)
            {
                foreach (var item in sheets)
                {
                    var sheetList = GetListFromWorksheet<T>(item, importSetting, excelName);
                    if (sheetList.Count > 0)
                    {
                        list.AddRange(sheetList);
                    }
                }
            }

            return list;
        }

        private static List<T> GetListFromWorksheet<T>(ExcelWorksheet sheet, ImportExcelSetting importSetting,string excelName) where T : ExcelBase, new()
        {
            List<T> list = new List<T>();
            if (sheet == null || sheet.Dimension == null)
                return list;

            Type entityType = typeof(T);
            List<string> errorMessages = new List<string>();

            int maxColumnNum = sheet.Dimension.End.Column;
            int maxRowNum = sheet.Dimension.End.Row;

            var mergeList = sheet.MergedCells;

            for (int m = 1; m <= maxColumnNum; m++)
            {
                int cellRow = importSetting.ImportStyle.DataRowIndex - 1;
                var cell = sheet.Cells[cellRow, m];
                if (cell.Merge)
                {
                    var index = sheet.GetMergeCellId(cellRow, m);
                    var range = sheet.MergedCells[index-1]?.Split(new char[1] {':' },StringSplitOptions.RemoveEmptyEntries);
                    if (range != null && range.Length > 0)
                    {
                        cell = sheet.Cells[range[0]];
                    }
                }

                var importColumn = importSetting.ImportColumns.Where(o => o.Text == cell.Text?.Trim()).FirstOrDefault();
                if (importColumn != null)
                {
                    importColumn.ColumnIndex = m;
                }
            }

            importSetting.ImportColumns = importSetting.ImportColumns.Where(o => o.ColumnIndex > 0).ToList();

            for (int n = importSetting.ImportStyle.DataRowIndex; n <= maxRowNum; n++)
            {
                var entity = Activator.CreateInstance<T>();
                //??
                entity.ExcelName = excelName;
                entity.SheetName = sheet.Name;
                entity.RowIndex = n;

                foreach (var item in importSetting.ImportColumns)
                {
                    entity.ColumnIndex = item.ColumnIndex;
                    var property = entityType.GetProperty(item.Name);

                    #region 数据校验
                    try
                    {
                        //??sheet.Cells[n, item.ColumnIndex].Text
                        var value = sheet.Cells[n, item.ColumnIndex].Value;
                        if (value != null && value.ToString() != "")
                        {
                            if(value is string)
                            {
                                value = ((string)value).TrimStart(item.Prefix.ToArray()).TrimEnd(item.Suffix.ToArray());
                            }

                            if (item.IsUseTempField)
                            {
                                var tempProperty = entityType.GetProperty(item.TempField);
                                if (tempProperty != null)
                                {
                                    tempProperty.SetValue(entity, value, null);
                                }
                                else
                                {
                                    var tempField = entityType.GetField(item.TempField);
                                    if (tempField != null)
                                    {
                                        tempField.SetValue(entity, value);
                                    }
                                }
                                continue;
                            }

                            if (!string.IsNullOrEmpty(item.DataVerReg))
                            {
                                var isMatch = Regex.IsMatch((string)value, item.DataVerReg);
                                if (!isMatch && item.IsRegFailThrowException)
                                {
                                    throw new ImportFormatException();
                                }
                            }

                            if (property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                            {
                                if (value != null && value.ToString().Length > 0)
                                {
                                    var tempValue = Convert.ChangeType(value, property.PropertyType.GetGenericArguments()[0]);
                                    property.SetValue(entity, tempValue, null);
                                }
                            }
                            else
                            {
                                var tempValue = Convert.ChangeType(value, property.PropertyType);
                                property.SetValue(entity, tempValue, null);
                            }
                        }

                    }
                    catch (ImportFormatException ex)
                    {
                        errorMessages.Add(string.Format("Excel{0}sheet:{1}第{2}行第{3}列【{4}】的值【{5}】不合法", entity.ExcelName, entity.SheetName, n, item.ColumnIndex, item.Text, sheet.Cells[n, item.ColumnIndex].Value));
                    }
                    catch (Exception ex)
                    {
                        errorMessages.Add(string.Format("Excel{0}sheet:{1}第{2}行第{3}列【{4}】的值【{5}】转换异常", entity.ExcelName, entity.SheetName, n, item.ColumnIndex, item.Text, sheet.Cells[n, item.ColumnIndex].Value));
                    }
                    #endregion
                }

                list.Add(entity);
            }

            if (errorMessages.Count > 0)
            {
                throw new ImportFormatException(errorMessages.ToArray());
            }

            return list;
        }
    }
}
