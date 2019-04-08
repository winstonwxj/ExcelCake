using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace ExcelCake.Intrusive
{
    public static class ExportExtension
    {
        private static readonly string _DateTimeFormatStr = "yyyyMMddHHmmssfff";
        private static readonly string _DefaultExcelName = "ExportFile";
        private static readonly string _ExportExcelNameTemplate = "{0}-{1}.xlsx";

        /// <summary>
        /// 导出List<T>为Stream数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static MemoryStream ExportToExcelStream<T>(this IEnumerable<T> list, string sheetName = "Sheet1") where T : ExcelBase,new()
        {
            //MemoryStream stream = new MemoryStream();
            //Type type = typeof(T);

            //var exportSetting = new ExportExcelSetting(type);

            //using (ExcelPackage package = new ExcelPackage())
            //{
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //    FillExcelWorksheet<T>(worksheet, list, exportSetting);
            //    package.SaveAs(stream);
            //}
            //return stream;
            var excelDic = new Dictionary<string, IEnumerable<ExcelBase>>();
            excelDic.Add(sheetName, list as IEnumerable<ExcelBase>);
            return ExportMultiToStream(excelDic);
        }

        /// <summary>
        /// 导出List<T>为byte[]数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static byte[] ExportToExcelBytes<T>(this IEnumerable<T> list, string sheetName = "Sheet1") where T : ExcelBase, new()
        {
            //byte[] excelBuffer = null;
            //Type type = typeof(T);

            //var exportSetting = new ExportExcelSetting(type);

            //using (ExcelPackage package = new ExcelPackage())
            //{
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //    FillExcelWorksheet<T>(worksheet, list, exportSetting);
            //    excelBuffer = package.GetAsByteArray();
            //}
            //return excelBuffer;
            var excelDic = new Dictionary<string, IEnumerable<ExcelBase>>();
            excelDic.Add(sheetName, list as IEnumerable<ExcelBase>);
            return ExportMultiToBytes(excelDic);
        }

        /// <summary>
        /// 导出List<T>为Excel文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheetName"></param>
        /// <param name="fileName"></param>
        public static string ExportToExcelFile<T>(this IEnumerable<T> list,string cacheDirectory, string sheetName = "Sheet1", string excelName = "") where T : ExcelBase, new()
        {
            //var dateTime = DateTime.Now.ToString(_dateTimeFormatStr);
            //if (excelName == "")
            //{
            //    excelName = _defaultExcelName;
            //}
            //if (excelName.IndexOf(".") > -1)
            //{
            //    excelName = excelName.Substring(0, excelName.IndexOf("."));
            //}

            //var fileName = string.Format(_exportExcelNameTemplate, excelName, dateTime);
            //DirectoryInfo dic = new DirectoryInfo(cacheDirectory);

            //if (!dic.Exists)
            //{
            //    dic.Create();
            //}

            //string downFilePath = Path.Combine(cacheDirectory, fileName);

            //Type type = typeof(T);
            //var exportSetting = new ExportExcelSetting(type);

            //using (ExcelPackage package = new ExcelPackage(new FileInfo(downFilePath)))
            //{
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //    FillExcelWorksheet<T>(worksheet, list, exportSetting);
            //    package.Save();
            //    return downFilePath;
            //}
            var excelDic = new Dictionary<string, IEnumerable<ExcelBase>>();
            excelDic.Add(sheetName, list as IEnumerable<ExcelBase>);
            return ExportMultiToFile(excelDic,cacheDirectory,excelName);
        }

        /// <summary>
        /// 导出List<T>为Stream数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dic"></param>
        /// <returns></returns>
        public static MemoryStream ExportMultiToStream(this IDictionary<string, IEnumerable<ExcelBase>> dic)
        {
            MemoryStream stream = new MemoryStream();
            
            using (ExcelPackage package = new ExcelPackage())
            {
                foreach(var item in dic)
                {
                    var types = item.Value.GetType().GetGenericArguments();
                    if (types != null && types.Length > 0)
                    {
                        var exportSetting = new ExportExcelSetting(types.First());
                        if (exportSetting.ExportColumns.Count > 0)
                        {
                            var worksheet = package.Workbook.Worksheets.Add(item.Key);
                            FillExcelWorksheet<ExcelBase>(worksheet, item.Value, exportSetting);
                        }
                    }
                }

                package.SaveAs(stream);
            }
            return stream;
        }

        /// <summary>
        /// 导出List<T>为byte[]数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dic"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static byte[] ExportMultiToBytes(this IDictionary<string, IEnumerable<ExcelBase>> dic)
        {
            byte[] excelBuffer = null;

            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (var item in dic)
                {
                    var types = item.Value.GetType().GetGenericArguments();
                    if (types != null && types.Length > 0)
                    {
                        var exportSetting = new ExportExcelSetting(types.First());
                        if (exportSetting.ExportColumns.Count > 0)
                        {
                            var worksheet = package.Workbook.Worksheets.Add(item.Key);
                            FillExcelWorksheet<ExcelBase>(worksheet, item.Value, exportSetting);
                        }
                    }
                }
                excelBuffer = package.GetAsByteArray();
            }
            return excelBuffer;
        }

        /// <summary>
        /// 导出List<T>为Excel文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheetName"></param>
        /// <param name="fileName"></param>
        public static string ExportMultiToFile(this IDictionary<string, IEnumerable<ExcelBase>> dic, string cacheDirectory,string excelName = "")
        {
            var dateTime = DateTime.Now.ToString(_DateTimeFormatStr);
            if (excelName == "")
            {
                excelName = _DefaultExcelName;
            }
            if (excelName.IndexOf(".") > -1)
            {
                excelName = excelName.Substring(0, excelName.IndexOf("."));
            }

            var fileName = string.Format(_ExportExcelNameTemplate, excelName, dateTime);
            DirectoryInfo dire = new DirectoryInfo(cacheDirectory);

            if (!dire.Exists)
            {
                dire.Create();
            }

            string downFilePath = Path.Combine(cacheDirectory, fileName);

            using (ExcelPackage package = new ExcelPackage(new FileInfo(downFilePath)))
            {
                foreach (var item in dic)
                {
                    var types = item.Value.GetType().GetGenericArguments();
                    if (types != null && types.Length > 0)
                    {
                        var exportSetting = new ExportExcelSetting(types.First());
                        if (exportSetting.ExportColumns.Count > 0)
                        {
                            var worksheet = package.Workbook.Worksheets.Add(item.Key);
                            FillExcelWorksheet<ExcelBase>(worksheet, item.Value, exportSetting);
                        }
                    }
                }
                package.Save();
                return downFilePath;
            }
        }

        private static void FillExcelWorksheet<T>(ExcelWorksheet sheet, IEnumerable<T> list, ExportExcelSetting exportSetting) where T : ExcelBase
        {
            if (sheet == null)
            {
                return;
            }
            Type type = null;
            var types = list.GetType().GetGenericArguments();
            if (types != null && types.Length > 0)
            {
                type = types.First();
            }
            else
            {
                type = typeof(T);
            }
            
            Color titleColor = Color.White;
            Color headColor = Color.White;
            Color contentColor = Color.White;

            int titleRowCount = 1;
            int startRow = 1;
            int startCol = 1;
            int endRow = 1;
            int endCol = 1;

            if (exportSetting.ExportStyle != null)
            {
                var title = exportSetting.ExportStyle.Title;
                endCol = (exportSetting.ExportColumns?.Count) ?? 1;

                if (exportSetting.ExportStyle.TitleColor != null)
                {
                    titleColor = exportSetting.ExportStyle.TitleColor;
                }

                if (exportSetting.ExportStyle.HeadColor != null)
                {
                    headColor = exportSetting.ExportStyle.HeadColor;
                }
                
                if (exportSetting.ExportStyle.ContentColor != null)
                {
                    contentColor = exportSetting.ExportStyle.ContentColor;
                }

                if (!string.IsNullOrEmpty(title) && endCol != 1)
                {
                    int titleEndCol = endCol;
                    if (exportSetting.ExportStyle.TitleColumnSpan > 0)
                    {
                        titleEndCol = exportSetting.ExportStyle.TitleColumnSpan;
                    }
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Merge = true;
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Value = title;
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.Font.Bold = exportSetting.ExportStyle.IsTitleBold;
                    if (exportSetting.ExportStyle.TitleFontSize > 0)
                    {
                        sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.Font.Size = exportSetting.ExportStyle.TitleFontSize;
                    }
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.Fill.BackgroundColor.SetColor(titleColor);
                    sheet.Cells[startRow, startCol, endRow, titleEndCol].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                }
                titleRowCount++;
            }

            //写入数据
            int mergeRowCount = 0;
            if(exportSetting.ExportColumns.GroupBy(o => o.MergeText).Any(o => o.Count() > 1&&o.First().MergeText!=null&&o.First().MergeText!=null))
            {
                mergeRowCount = 1;
            }
            int dataStartRow = titleRowCount+ mergeRowCount;
            int dataStartCol = 1;

            for (var i = 0; i < exportSetting.ExportColumns.Count; i++)
            {
                if (mergeRowCount >= 1 && string.IsNullOrEmpty(exportSetting.ExportColumns[i].MergeText))
                {
                    sheet.Cells[dataStartRow - 1, i + dataStartCol, dataStartRow, i + dataStartCol].Merge = true;
                    sheet.Cells[dataStartRow - 1, i + dataStartCol].Value = exportSetting.ExportColumns[i].Text;
                    sheet.Cells[dataStartRow - 1, i + dataStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[dataStartRow - 1, i + dataStartCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[dataStartRow - 1, i + dataStartCol].Style.Font.Bold = exportSetting.ExportStyle.IsHeadBold;
                    if (exportSetting.ExportStyle.HeadFontSize > 0)
                    {
                        sheet.Cells[dataStartRow - 1, i + dataStartCol].Style.Font.Size = exportSetting.ExportStyle.HeadFontSize;
                    }
                    sheet.Cells[dataStartRow - 1, i + dataStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[dataStartRow - 1, i + dataStartCol].Style.Fill.BackgroundColor.SetColor(headColor);
                    sheet.Cells[dataStartRow - 1, i + dataStartCol, dataStartRow, i + dataStartCol].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }
                else if (mergeRowCount >= 1 && !string.IsNullOrEmpty(exportSetting.ExportColumns[i].MergeText))
                {
                    var mergeList = exportSetting.MergeList.Where(o => o.Key == exportSetting.ExportColumns[i].MergeText)?.ToList();

                    if (mergeList != null && mergeList.Count > 0)
                    {
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol, dataStartRow - mergeRowCount, i + dataStartCol + mergeList.First().Value - 1].Merge = true;
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Value = exportSetting.ExportColumns[i].MergeText;
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Style.Font.Bold = exportSetting.ExportStyle.IsHeadBold;
                        if (exportSetting.ExportStyle.HeadFontSize > 0)
                        {
                            sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Style.Font.Size = exportSetting.ExportStyle.HeadFontSize;
                        }

                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid; ;
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol].Style.Fill.BackgroundColor.SetColor(headColor);
                        sheet.Cells[dataStartRow - mergeRowCount, i + dataStartCol, dataStartRow - mergeRowCount, i + dataStartCol + mergeList.First().Value - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        exportSetting.MergeList.Remove(mergeList.First());
                    }

                    sheet.Cells[dataStartRow, i + dataStartCol].Value = exportSetting.ExportColumns[i].Text;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Font.Bold = exportSetting.ExportStyle.IsHeadBold;
                    if (exportSetting.ExportStyle.HeadFontSize > 0)
                    {
                        sheet.Cells[dataStartRow, i + dataStartCol].Style.Font.Size = exportSetting.ExportStyle.HeadFontSize;
                    }
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Fill.BackgroundColor.SetColor(headColor);
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }
                else
                {
                    sheet.Cells[dataStartRow, i + dataStartCol].Value = exportSetting.ExportColumns[i].Text;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Font.Bold = exportSetting.ExportStyle.IsHeadBold;
                    if (exportSetting.ExportStyle.HeadFontSize > 0)
                    {
                        sheet.Cells[dataStartRow, i + dataStartCol].Style.Font.Size = exportSetting.ExportStyle.HeadFontSize;
                    }
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Fill.BackgroundColor.SetColor(headColor);
                    sheet.Cells[dataStartRow, i + dataStartCol].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }

                int j = 0;
                foreach (var item in list)
                {
                    object value = null;
                    PropertyInfo propertyInfo = type.GetProperty(exportSetting.ExportColumns[i].Value);
                    value = propertyInfo.GetValue(item, null);
                    if(value!=null&&value is string)
                    {
                        value = ((string)value).TrimStart(exportSetting.ExportColumns[i].Prefix.ToArray()).TrimEnd(exportSetting.ExportColumns[i].Suffix.ToArray());
                    }
                    //try
                    //{
                    //    PropertyInfo propertyInfo = type.GetProperty(exportSetting.ExportColumns[i].Value);
                    //    value = propertyInfo.GetValue(item, null);
                    //}
                    //catch (Exception ex)
                    //{
                    //    value = "";
                    //}

                    if (value != null && value.ToString()!="")
                    {
                        value = exportSetting.ExportColumns[i].Prefix+value+exportSetting.ExportColumns[i].Suffix;
                    }

                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Value = value??"";
                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.Font.Bold = exportSetting.ExportStyle.IsContentBold;
                    if (exportSetting.ExportStyle.ContentFontSize > 0)
                    {
                        sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.Font.Size = exportSetting.ExportStyle.ContentFontSize;
                    }
                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.Fill.BackgroundColor.SetColor(contentColor);
                    sheet.Cells[j + dataStartRow + 1, i + dataStartCol].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    j++;
                }
                sheet.Column(i + 1).AutoFit();
            }
        }
    }
}