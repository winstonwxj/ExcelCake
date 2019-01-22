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

namespace XLSXCake.Intrusive
{
    public static class ListExtension
    {
        /// <summary>
        /// 导出List<T>为Stream数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static MemoryStream ExportToExcelStream<T>(this List<T> list, string sheetName = "Sheet1") where T : XLSXBase,new()
        {
            MemoryStream stream = new MemoryStream();
            Type type = typeof(T);

            ExportExcelSetting exportSetting = GetExportColumns(type);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = FillExcelWorksheet<T>(package, exportSetting, list, sheetName: sheetName);
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
        public static byte[] ExportToExcelBytes<T>(this List<T> list, string sheetName = "Sheet1") where T : XLSXBase, new()
        {
            byte[] excelBuffer = null;
            Type type = typeof(T);

            ExportExcelSetting exportSetting = GetExportColumns(type);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = FillExcelWorksheet<T>(package, exportSetting, list, sheetName: sheetName);
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
        public static string ExportToExcelFile<T>(this List<T> list,string cacheDirectory, string sheetName = "Sheet1", string excelName = "") where T : XLSXBase, new()
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

            ExportExcelSetting exportSetting = GetExportColumns(type);

            using (ExcelPackage package = new ExcelPackage(new FileInfo(downFilePath)))
            {
                ExcelWorksheet worksheet = FillExcelWorksheet<T>(package, exportSetting, list, sheetName: sheetName);
                package.Save();
                return downFilePath;
            }
        }

        private static ExcelWorksheet FillExcelWorksheet<T>(ExcelPackage package, ExportExcelSetting exportSetting, List<T> list, string sheetName = "Sheet1") where T : XLSXBase, new()
        {
            Type type = typeof(T);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //Color headColor = Color.FromArgb(192, 192, 192);
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
                    worksheet.Cells[1, 1, count, count].Merge = true;
                    worksheet.Cells[1, 1, count, count].Value = title;
                    worksheet.Cells[1, 1, count, count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[1, 1, count, count].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[1, 1, count, count].Style.Font.Bold = true;
                    worksheet.Cells[1, 1, count, count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //worksheet.Cells[1, 1, count, count].Style.Fill.BackgroundColor.SetColor(headColor);
                    //worksheet.Cells[1, 1, count, count].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                }
                columnIndex++;
            }

            //写入数据
            for (var i = 0; i < exportSetting.ExportColumns.Count; i++)
            {
                worksheet.Cells[columnIndex, i + 1].Value = exportSetting.ExportColumns[i].Text;
                worksheet.Cells[columnIndex, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[columnIndex, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[columnIndex, i + 1].Style.Font.Bold = true;
                worksheet.Cells[columnIndex, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
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
                    worksheet.Cells[j + columnIndex + 2, i + 1].Value = value ?? "";
                    worksheet.Cells[j + columnIndex + 2, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[j + columnIndex + 2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //worksheet.Cells[j + columnIndex + 2, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }
                worksheet.Column(i + 1).AutoFit();
            }
            return worksheet;
        }

        private static ExportExcelSetting GetExportColumns(Type type)
        {
            ExportExcelSetting exportSetting = new ExportExcelSetting();
            
            //List<ExportColumn> columns = new List<ExportColumn>();
            #region 组织表头
            //判断类导出模式
            var modalType = EnumExportModalType.ALL;
            var classAttrArry = type.GetCustomAttributes(typeof(ExportModalAttribute), true);
            if (classAttrArry != null && classAttrArry.Length > 0)
            {
                modalType = ((ExportModalAttribute)classAttrArry[0]).ExportModal;
                exportSetting.ExportStyle.Title = ((ExportModalAttribute)classAttrArry[0]).Title;
                //exportSetting.ExportStyle.HeadColor = ((ExportModalAttribute)classAttrArry[0]).HeadColor;
            }
            
            //导出字段
            var properties = type.GetProperties();
            if (modalType == EnumExportModalType.PART)
            {//部分模式
                foreach (var proper in properties)
                {
                    var exportAttrArry = proper.GetCustomAttributes(typeof(ExportAttribute), true);
                    if (exportAttrArry != null && exportAttrArry.Length > 0)
                    {
                        string displayName = proper.Name;
                        var nameAttrArry = proper.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                        if (nameAttrArry != null && nameAttrArry.Length > 0)
                        {
                            displayName = ((DisplayNameAttribute)nameAttrArry[0]).DisplayName;
                        }
                        int index = 0;
                        var indexAttrArry = proper.GetCustomAttributes(typeof(ExportSortAttribute), true);
                        if (indexAttrArry != null && indexAttrArry.Length > 0)
                        {
                            index = ((ExportSortAttribute)indexAttrArry[0]).SortIndex;
                        }
                        exportSetting.ExportColumns.Add(new ExportColumn()
                        {
                            Text = displayName,
                            Value = proper.Name,
                            Index = index
                        });
                    }

                }
            }
            else if (modalType == EnumExportModalType.INVERT)
            {//反选模式
                foreach (var proper in properties)
                {
                    var noexportAttrArry = proper.GetCustomAttributes(typeof(NoExportAttribute), true);
                    if (noexportAttrArry == null || noexportAttrArry.Length == 0)
                    {
                        string displayName = proper.Name;
                        var nameAttrArry = proper.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                        if (nameAttrArry != null && nameAttrArry.Length > 0)
                        {
                            displayName = ((DisplayNameAttribute)nameAttrArry[0]).DisplayName;
                        }
                        int index = 0;
                        var indexAttrArry = proper.GetCustomAttributes(typeof(ExportSortAttribute), true);
                        if (indexAttrArry != null && indexAttrArry.Length > 0)
                        {
                            index = ((ExportSortAttribute)indexAttrArry[0]).SortIndex;
                        }
                        exportSetting.ExportColumns.Add(new ExportColumn()
                        {
                            Text = displayName,
                            Value = proper.Name,
                            Index = index
                        });
                    }
                }
            }
            else
            {//全选模式
                foreach (var proper in properties)
                {
                    string displayName = proper.Name;
                    var nameAttrArry = proper.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                    if (nameAttrArry != null && nameAttrArry.Length > 0)
                    {
                        displayName = ((DisplayNameAttribute)nameAttrArry[0]).DisplayName;
                    }
                    int index = 0;
                    var indexAttrArry = proper.GetCustomAttributes(typeof(ExportSortAttribute), true);
                    if (indexAttrArry != null && indexAttrArry.Length > 0)
                    {
                        index = ((ExportSortAttribute)indexAttrArry[0]).SortIndex;
                    }
                    exportSetting.ExportColumns.Add(new ExportColumn()
                    {
                        Text = displayName,
                        Value = proper.Name,
                        Index = index
                    });
                }
            }
            #endregion

            #region 排序
            exportSetting.ExportColumns.Sort((a, b) => a.Index.CompareTo(b.Index));
            #endregion
            return exportSetting;
        }

        private static DataTable GetExportDataTable<T>(List<T> list, List<ExportColumn> columns) where T : XLSXBase, new()
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dt = new DataTable();
            if (columns == null || columns.Count == 0)
            {
                throw new Exception("导出异常");
            }

            foreach (var item in columns)
            {
                dt.Columns.Add(item.Value, typeof(string));
            }

            if (list != null)
            {
                foreach (var item in list)
                {
                    var row = dt.NewRow();
                    foreach (var column in columns)
                    {
                        string value = "";
                        try
                        {
                            PropertyInfo propertyInfo = type.GetProperty(column.Value);
                            value = (string)propertyInfo.GetValue(item, null);
                        }
                        catch (Exception ex)
                        {

                        }
                        row[column.Value] = value;
                    }
                    dt.Rows.Add(row);
                }
            }

            return dt;
        }
    }

    internal class ExportColumn
    {
        public string Text { set; get; }
        public string Value { set; get; }
        public int Index { set; get; }
    }

    internal class ExportStyle
    {
        public string Title { set; get; }

        //public Color HeadColor { set; get; }
    }

    internal class ExportExcelSetting
    {
        public List<ExportColumn> ExportColumns { set; get; }
        public ExportStyle ExportStyle { set; get; }

        public ExportExcelSetting()
        {
            ExportStyle = new ExportStyle();
            ExportColumns = new List<ExportColumn>();
        }
    }
}