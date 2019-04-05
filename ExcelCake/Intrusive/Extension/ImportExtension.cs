using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelCake.Intrusive
{
    public static class ImportExtension
    {
        private static readonly string _DateTimeFormatStr = "yyyyMMddHHmmssfff";

        public static void ImportToList<T>(this IEnumerable<T> list, FileInfo file, List<string> importSheets=null,List<string> noImportSheets=null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {

            if (!file.Exists)
            {
                list = new List<T>();
                return;
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

                list = GetCollectionFromSheets<T>(sheets);
            }
        }

        public static void ImportToList<T>(this IEnumerable<T> list, string filePath, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            list.ImportToList(new FileInfo(filePath), importSheets, noImportSheets, importSheetsRegex, noImportSheetsRegex);
        }

        public static void ImportToAppendList<T>(this IEnumerable<T> list, FileInfo file, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            List<T> tempList = list?.ToList();
            List<T> tempList2 = new List<T>();
            if (file.Exists)
            {
                tempList2.ImportToList<T>(file, importSheets, noImportSheets, importSheetsRegex, noImportSheetsRegex);
                if (tempList2.Count > 0)
                {
                    tempList.AddRange(tempList2);
                }
            }
            
            list = tempList;
        }

        public static void ImportToAppendList<T>(this IEnumerable<T> list, string filePath, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            list.ImportToAppendList(new FileInfo(filePath), importSheets, noImportSheets, importSheetsRegex, noImportSheetsRegex);
        }

        private static IEnumerable<T> GetCollectionFromSheets<T>(List<ExcelWorksheet> sheets) where T : ExcelBase, new()
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
                foreach(var item in sheets)
                {
                    var sheetList = GetListFromWorksheet<T>(item, importSetting);
                    if (sheetList.Count>0)
                    {
                        list.AddRange(sheetList);
                    }
                }
            }

            return list;
        }
        
        private static List<T> GetListFromWorksheet<T>(ExcelWorksheet sheet,ImportExcelSetting importSetting) where T : ExcelBase, new()
        {
            List<T> list = new List<T>();
            if (sheet == null||sheet.Dimension == null)
                return list;

            Type entityType = typeof(T);
            List<string> errorMessages = new List<string>();

            int maxColumnNum = sheet.Dimension.End.Column;
            int maxRowNum = sheet.Dimension.End.Row;

            for (int m = 1; m <= maxColumnNum; m++)
            {
                var cell = sheet.Cells[importSetting.ImportStyle.HeadRowIndex, m];
                var importColumn = importSetting.ImportColumns.Where(o => o.Text == cell.Text?.Trim()).FirstOrDefault();
                if (importColumn != null)
                {
                    importColumn.ColumnIndex = m;
                }
            }

            for (int n = importSetting.ImportStyle.DataRowIndex; n <= maxRowNum; n++)
            {
                var entity = Activator.CreateInstance<T>();
                //??
                entity.ExcelName = sheet.Workbook.Properties.Title;
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
                            if (item.IsConvert)
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
                    catch (Exception ex)
                    {
                        errorMessages.Add(string.Format("Excel{0}sheet:{1}第{2}行第{3}列【{4}】的值【{5}】不合法",entity.ExcelName,entity.SheetName, n, item.ColumnIndex, item.Text, sheet.Cells[n, item.ColumnIndex].Value));
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
