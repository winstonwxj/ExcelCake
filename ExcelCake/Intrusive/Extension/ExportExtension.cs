using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace ExcelCake.Intrusive
{
    public static class ExportExtension
    {
        /// <summary>
        /// 导出List<T>为Stream数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static MemoryStream ExportToExcelStream<T>(this List<T> list, string sheetName = "Sheet1") where T : ExcelBase,new()
        {
            MemoryStream stream = new MemoryStream();
            Type type = typeof(T);

            var exportSetting = new ExportExcelSetting(type);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                FillExcelWorksheet<T>(worksheet, list, exportSetting);
                package.SaveAs(stream);
            }
            return stream;
        }

        /// <summary>
        /// 导出List<T>为byte[]数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static byte[] ExportToExcelBytes<T>(this List<T> list, string sheetName = "Sheet1") where T : ExcelBase, new()
        {
            byte[] excelBuffer = null;
            Type type = typeof(T);

            var exportSetting = new ExportExcelSetting(type);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                FillExcelWorksheet<T>(worksheet, list, exportSetting);
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
        public static string ExportToExcelFile<T>(this List<T> list,string cacheDirectory, string sheetName = "Sheet1", string excelName = "") where T : ExcelBase, new()
        {
            var dateTime = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            if (excelName == "")
            {
                excelName = "ExportFile";
            }
            if (excelName.IndexOf(".") > -1)
            {
                excelName = excelName.Substring(0, excelName.IndexOf("."));
            }

            var fileName = string.Format("{0}-{1}.xlsx", excelName, dateTime);
            DirectoryInfo dic = new DirectoryInfo(cacheDirectory);

            if (!dic.Exists)
            {
                dic.Create();
            }

            string downFilePath = Path.Combine(cacheDirectory, fileName);

            Type type = typeof(T);
            var exportSetting = new ExportExcelSetting(type);

            using (ExcelPackage package = new ExcelPackage(new FileInfo(downFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                FillExcelWorksheet<T>(worksheet, list, exportSetting);
                package.Save();
                return downFilePath;
            }
        }

        private static void FillExcelWorksheet<T>(ExcelWorksheet sheet, List<T> list, ExportExcelSetting exportSetting) where T : ExcelBase
        {
            if (sheet == null)
            {
                return;
            }
            Type type = typeof(T);
            //ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

            int columnIndex = 1;
            if (exportSetting.ExportStyle != null)
            {
                var title = exportSetting.ExportStyle.Title;
                var count = exportSetting.ExportColumns!=null?exportSetting.ExportColumns.Count:0;
                //if (exportSetting.ExportStyle.HeadColor != null)
                //{
                //    headColor = exportSetting.ExportStyle.HeadColor;
                //}
                if (!string.IsNullOrEmpty(title) && count != 0)
                {
                    sheet.Cells[1, 1, count, count].Merge = true;
                    sheet.Cells[1, 1, count, count].Value = title;
                    sheet.Cells[1, 1, count, count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[1, 1, count, count].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[1, 1, count, count].Style.Font.Bold = true;
                    sheet.Cells[1, 1, count, count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //worksheet.Cells[1, 1, count, count].Style.Fill.BackgroundColor.SetColor(headColor);
                    //worksheet.Cells[1, 1, count, count].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                }
                columnIndex++;
            }

            //写入数据
            for (var i = 0; i < exportSetting.ExportColumns.Count; i++)
            {
                sheet.Cells[columnIndex, i + 1].Value = exportSetting.ExportColumns[i].Text;
                sheet.Cells[columnIndex, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[columnIndex, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[columnIndex, i + 1].Style.Font.Bold = true;
                sheet.Cells[columnIndex, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[columnIndex, i + 1].Style.Fill.BackgroundColor.SetColor(headColor);
                //worksheet.Cells[columnIndex, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                for (var j = 0; j < list.Count; j++)
                {
                    object value = null;
                    try
                    {
                        PropertyInfo propertyInfo = type.GetProperty(exportSetting.ExportColumns[i].Value);
                        value = propertyInfo.GetValue(list[j], null);
                    }
                    catch (Exception ex)
                    {
                        value = "";
                    }
                    sheet.Cells[j + columnIndex + 2, i + 1].Value = value ?? "";
                    sheet.Cells[j + columnIndex + 2, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[j + columnIndex + 2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //worksheet.Cells[j + columnIndex + 2, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }
                sheet.Column(i + 1).AutoFit();
            }
        }

        //private static ExcelWorksheet FillExcelWorksheet<T>(ExcelPackage package, ExportExcelSetting exportSetting, List<T> list, string sheetName = "Sheet1") where T : ExcelBase, new()
        //{
        //    Type type = typeof(T);
        //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
        //    //Color headColor = Color.FromArgb(192, 192, 192);
        //    int columnIndex = 1;
        //    if (exportSetting.ExportStyle != null)
        //    {
        //        var title = exportSetting.ExportStyle.Title;
        //        var count = exportSetting.ExportColumns != null ? exportSetting.ExportColumns.Count : 0;
        //        //if (exportSetting.ExportStyle.HeadColor != null)
        //        //{
        //        //    headColor = exportSetting.ExportStyle.HeadColor;
        //        //}
        //        if (!string.IsNullOrEmpty(title) && count != 0)
        //        {
        //            worksheet.Cells[1, 1, count, count].Merge = true;
        //            worksheet.Cells[1, 1, count, count].Value = title;
        //            worksheet.Cells[1, 1, count, count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[1, 1, count, count].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[1, 1, count, count].Style.Font.Bold = true;
        //            worksheet.Cells[1, 1, count, count].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //            //worksheet.Cells[1, 1, count, count].Style.Fill.BackgroundColor.SetColor(headColor);
        //            //worksheet.Cells[1, 1, count, count].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

        //        }
        //        columnIndex++;
        //    }

        //    //写入数据
        //    for (var i = 0; i < exportSetting.ExportColumns.Count; i++)
        //    {
        //        worksheet.Cells[columnIndex, i + 1].Value = exportSetting.ExportColumns[i].Text;
        //        worksheet.Cells[columnIndex, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //        worksheet.Cells[columnIndex, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //        worksheet.Cells[columnIndex, i + 1].Style.Font.Bold = true;
        //        worksheet.Cells[columnIndex, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        //worksheet.Cells[columnIndex, i + 1].Style.Fill.BackgroundColor.SetColor(headColor);
        //        //worksheet.Cells[columnIndex, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //        for (var j = 0; j < list.Count; j++)
        //        {
        //            object value = null;
        //            try
        //            {
        //                PropertyInfo propertyInfo = type.GetProperty(exportSetting.ExportColumns[i].Value);
        //                value = propertyInfo.GetValue(list[j], null);
        //            }
        //            catch (Exception ex)
        //            {
        //                value = "";
        //            }
        //            worksheet.Cells[j + columnIndex + 2, i + 1].Value = value ?? "";
        //            worksheet.Cells[j + columnIndex + 2, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[j + columnIndex + 2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            //worksheet.Cells[j + columnIndex + 2, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //        }
        //        worksheet.Column(i + 1).AutoFit();
        //    }
        //    return worksheet;
        //}

        /// <summary>
        /// 导出List<T>为Stream数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dic"></param>
        /// <returns></returns>
        public static MemoryStream ExportToExcelStream(this Dictionary<string, List<ExcelBase>> dic)
        {
            MemoryStream stream = new MemoryStream();
            
            using (ExcelPackage package = new ExcelPackage())
            {
                foreach(var item in dic)
                {
                    Type type = item.Value.GetType();
                    var exportSetting = new ExportExcelSetting(type);
                    var worksheet = package.Workbook.Worksheets.Add(item.Key);
                    FillExcelWorksheet<ExcelBase>(worksheet, item.Value, exportSetting);
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
        public static byte[] ExportToExcelBytes<T>(this Dictionary<string, List<ExcelBase>> dic)
        {
            byte[] excelBuffer = null;

            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (var item in dic)
                {
                    Type type = item.Value.GetType();
                    var exportSetting = new ExportExcelSetting(type);
                    var worksheet = package.Workbook.Worksheets.Add(item.Key);
                    FillExcelWorksheet<ExcelBase>(worksheet, item.Value, exportSetting);
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
        public static string ExportToExcelFile<T>(this Dictionary<string, List<ExcelBase>> dic, string cacheDirectory,string excelName = "")
        {
            var dateTime = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            if (excelName == "")
            {
                excelName = "ExportFile";
            }
            if (excelName.IndexOf(".") > -1)
            {
                excelName = excelName.Substring(0, excelName.IndexOf("."));
            }

            var fileName = string.Format("{0}-{1}.xlsx", excelName, dateTime);
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
                    Type type = item.Value.GetType();
                    var exportSetting = new ExportExcelSetting(type);
                    var worksheet = package.Workbook.Worksheets.Add(item.Key);
                    FillExcelWorksheet<ExcelBase>(worksheet, item.Value, exportSetting);
                }
                package.Save();
                return downFilePath;
            }
        }
    }
}