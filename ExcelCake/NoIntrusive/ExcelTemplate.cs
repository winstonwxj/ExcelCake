using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using System.Configuration;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelCake.NoIntrusive
{
    public class TemplateSetting
    {
        public string Type { set; get; }
        public string DataSource { set; get; }
        //public string Address { set; get; }
        public string AddressLeftTop { set; get; }
        public string AddressRightBottom { set; get; }
        public string Field { set; get; }
        public string SettingString { set; get; }
        public ExcelRangeBase CurrentCell { set; get; }
        public int FromRow { set; get; }
        public int FromCol { set; get; }
        public int ToRow { set; get; }
        public int ToCol { set; get; }
    }

    /// <summary>
    /// 自定义模板复杂格式导出(待重构)
    /// </summary>
    public class ExcelTemplate
    {
        private string _TemplateFile;
        private string _TemplateSheetName;

        private ExcelTemplate()
        {

        }

        public ExcelTemplate(string templateFilePath,string sheetName="Sheet1")
        {
            _TemplateFile = templateFilePath;
            _TemplateSheetName = sheetName;
        }

        /// <summary>
        /// 填充报表(待重构)
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="dataSource"></param>
        /// <returns></returns>
        public ExcelWorksheet FillSheetData(ExcelWorksheet workSheet,DataObject dataSource)
        {
            if (workSheet == null || dataSource == null)
            {
                return workSheet;
            }

            #region 分析配置
            List<TemplateSetting> gridSettingList = new List<TemplateSetting>();
            List<TemplateSetting> freeSettingList = new List<TemplateSetting>();
            List<TemplateSetting> fieldList = new List<TemplateSetting>();
            AnalysisSettings(workSheet, out freeSettingList, out gridSettingList, out fieldList);
            #endregion

            #region 填充数据
            foreach (var item in freeSettingList)
            {
                if (string.IsNullOrEmpty(item.AddressLeftTop) || string.IsNullOrEmpty(item.AddressRightBottom))
                {
                    continue;
                }
                ExcelRange range = workSheet.Cells[item.AddressLeftTop+","+item.AddressRightBottom];

                if (!dataSource.DataEntity.Keys.Contains(item.DataSource))
                {
                    continue;
                }
                var data = dataSource.DataEntity[item.DataSource];

                foreach(var field in fieldList)
                {
                    if (IsCellInRange(item.FromRow,item.FromCol,item.ToRow,item.ToCol,field.CurrentCell))
                    {
                        var fieldName = field.Field;
                        var value = field.CurrentCell.Value?.ToString() ?? "";
                        value = value.Replace(field.SettingString, "");
                        object fieldData = null;
                        if (data.Keys.Contains(field.Field))
                        {
                            fieldData = data[field.Field];
                        }
                        field.CurrentCell.Value = value + fieldData;
                    }
                }

            }

            List<string> mergedList = new List<string>();
            foreach(var item in workSheet.MergedCells)
            {
                mergedList.Add(item.Replace(":",","));
            }
            Dictionary<TemplateSetting, int> regionAddDic = new Dictionary<TemplateSetting, int>();

            foreach (var item in gridSettingList)
            {
                if (string.IsNullOrEmpty(item.AddressLeftTop) || string.IsNullOrEmpty(item.AddressRightBottom))
                {
                    continue;
                }
                ExcelRange range = workSheet.Cells[item.AddressLeftTop + "," + item.AddressRightBottom];

                if (!dataSource.DataList.Keys.Contains(item.DataSource))
                {
                    continue;
                }
                var data = dataSource.DataList[item.DataSource];

                //分析
                int offsetCount = 0;
                int addCount = 0;
                int sameCount = 0;
                int emptyCount = 0;

                if (data.Count > 0 && data.First().Value.Count() > 1)
                {
                    addCount = data.First().Value.Count() - 1;
                }
                foreach (var addItem in regionAddDic)
                {
                    if (addItem.Key.FromRow < item.FromRow&&addItem.Value>offsetCount)
                    {
                        offsetCount = addItem.Value;
                    }
                    else if(addItem.Key.FromRow == item.FromRow)
                    {
                        emptyCount = addItem.Value > addCount ? addItem.Value - addCount : 0;
                        addCount = addItem.Value>addCount?0:addCount-addItem.Value;
                        sameCount = addItem.Value > addCount ? 0 : addItem.Value;
                    }
                }

                //动态添加行
                if (addCount>0)
                {
                    workSheet.InsertRow(item.FromRow+offsetCount + 1, addCount);
                    regionAddDic.Add(item, addCount + sameCount);
                }
                
                foreach (var field in fieldList)
                {
                    if (IsCellInRange(item.FromRow, item.FromCol, item.ToRow, item.ToCol, field.CurrentCell))
                    {
                        var fieldName = field.Field;
                        List<object> fieldDatas = new List<object>();
                        
                        if (data.Keys.Contains(field.Field))
                        {
                            bool isFieldMerge = false;
                            int fromRow = 1;
                            int fromCol = 1;
                            int toRow = 1;
                            int toCol = 1;

                            foreach (var merge in mergedList)
                            {
                                var arryMerge = merge.Split(new char[1] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                if (arryMerge.Length == 2)
                                {
                                    var isMerge = IsCellInRange(arryMerge[0], arryMerge[1], field.CurrentCell);
                                    if (isMerge)
                                    {
                                        isFieldMerge = true;
                                        CalcRowCol(arryMerge[0], out fromRow, out fromCol);
                                        CalcRowCol(arryMerge[1], out toRow, out toCol);
                                        break;
                                    }
                                }
                            }

                            fieldDatas = data[field.Field];
                            
                            for (var i = 0; i < fieldDatas.Count; i++)
                            {
                                field.CurrentCell.Offset(i+ offsetCount, 0).Value = fieldDatas[i];
                                if (i == 0)
                                {
                                    continue;
                                }
                                
                                if (isFieldMerge)
                                {
                                    var oldRange = workSheet.Cells[fromRow+ offsetCount, fromCol, toRow+ offsetCount, toCol];
                                    var newRange = workSheet.Cells[fromRow+ offsetCount + i, fromCol, toRow+ offsetCount + i, toCol];
                                    newRange.Merge = true;

                                    HandleCellStyle(oldRange, newRange);
                                }
                                else
                                {
                                    var oldRange = field.CurrentCell.Offset(offsetCount, 0);
                                    var newRange = field.CurrentCell.Offset(i + offsetCount, 0);

                                    HandleCellStyle(oldRange, newRange);
                                }
                            }

                            if (emptyCount > 0)
                            {
                                for (var i = 0; i < emptyCount; i++)
                                {
                                    if (isFieldMerge)
                                    {
                                        var oldRange = workSheet.Cells[fromRow + offsetCount, fromCol, toRow + offsetCount, toCol];
                                        var newRange = workSheet.Cells[fromRow + offsetCount + fieldDatas.Count + i, fromCol, toRow + offsetCount + fieldDatas.Count + i, toCol];
                                        newRange.Merge = true;
                                        HandleCellStyle(oldRange, newRange);
                                    }
                                    else
                                    {
                                        var oldRange = field.CurrentCell.Offset(offsetCount, 0);
                                        var newRange = field.CurrentCell.Offset(i + offsetCount+ fieldDatas.Count, 0);
                                        HandleCellStyle(oldRange, newRange);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            return workSheet;
        }

        /// <summary>
        /// 分析配置
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="freeSettingList"></param>
        /// <param name="gridSettingList"></param>
        /// <param name="fieldSettingList"></param>
        public void AnalysisSettings(ExcelWorksheet sheet,out List<TemplateSetting> freeSettingList,out List<TemplateSetting> gridSettingList,out List<TemplateSetting> fieldSettingList)
        {
            gridSettingList = new List<TemplateSetting>();
            freeSettingList = new List<TemplateSetting>();
            fieldSettingList = new List<TemplateSetting>();
            if (sheet == null || sheet.Cells.Count() <= 0)
            {
                return;
            }
            foreach (var cell in sheet.Cells)
            {
                var cellValue = cell.Value?.ToString() ?? "";
                var arry = cellValue.Split(new char[2] { '{', '}' }, StringSplitOptions.RemoveEmptyEntries);
                if (arry.Length == 0)
                {
                    continue;
                }
                foreach (var item in arry)
                {
                    if (item.IndexOf(":") > -1 && item.IndexOf(";") > -1)
                    {
                        var settingItemArry = item.Split(new char[1] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        if (settingItemArry.Length == 0)
                        {
                            continue;
                        }
                        var setting = new TemplateSetting();
                        foreach (var arryItem in settingItemArry)
                        {
                            var settingItem = arryItem.Split(':');
                            if (settingItem.Length < 2)
                            {
                                continue;
                            }
                            var key = settingItem[0];
                            var value = settingItem[1];
                            if (string.IsNullOrEmpty(key))
                            {
                                continue;
                            }

                            switch (key.ToUpper())
                            {
                                case "TYPE": { setting.Type = value.ToUpper(); } break;
                                case "DATASOURCE": { setting.DataSource = value.ToUpper(); } break;
                                case "ADDRESSLEFTTOP": { setting.AddressLeftTop = value.ToUpper(); } break;
                                case "ADDRESSRIGHTBOTTOM": { setting.AddressRightBottom = value.ToUpper(); } break;
                                case "FIELD": { setting.Field = value.ToUpper(); } break;
                            }
                        }
                        setting.CurrentCell = cell;
                        setting.SettingString = "{" + item + "}";
                        if (string.IsNullOrEmpty(setting.Type))
                        {
                            continue;
                        }
                        else if (setting.Type == "GRID")
                        {
                            CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                            CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                            setting.FromRow = fromRow;
                            setting.FromCol = fromCol;
                            setting.ToRow = toRow;
                            setting.ToCol = toCol;
                            gridSettingList.Add(setting);
                        }
                        else if (setting.Type == "FREE")
                        {
                            CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                            CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                            setting.FromRow = fromRow;
                            setting.FromCol = fromCol;
                            setting.ToRow = toRow;
                            setting.ToCol = toCol;
                            freeSettingList.Add(setting);
                        }
                        else if (setting.Type == "VALUE")
                        {
                            fieldSettingList.Add(setting);
                        }

                    }
                    var cellValueStr = cell.Value?.ToString() ?? "";
                    cell.Value = cellValueStr.Replace("{" + item + "}", "");
                }
            }
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="settingCell"></param>
        /// <param name="newCell"></param>
        public void HandleCellStyle(ExcelRangeBase settingCell,ExcelRangeBase newCell)
        {
            if (settingCell == null || newCell == null)
            {
                return;
            }
            newCell.Style.HorizontalAlignment = settingCell.Style.HorizontalAlignment;
            newCell.Style.VerticalAlignment = settingCell.Style.VerticalAlignment;

            //
            newCell.Style.Font.Bold = settingCell.Style.Font.Bold;//字体为粗体
                                                                  //field.CurrentCell.Offset(i, 0).Style.Font.Color.SetColor();//字体颜色
            newCell.Style.Font.Name = settingCell.Style.Font.Name;//字体
            newCell.Style.Font.Size = settingCell.Style.Font.Size;//字体大小
            newCell.Style.Fill.PatternType = settingCell.Style.Fill.PatternType;
            //field.CurrentCell.Offset(i, 0).Style.Fill.BackgroundColor.SetColor();//设置单元格背
            //field.CurrentCell.Offset(i, 0).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));//设置单元格所有边框
            newCell.Style.Border.Left.Style = settingCell.Style.Border.Left.Style;
            newCell.Style.Border.Right.Style = settingCell.Style.Border.Right.Style;
            newCell.Style.Border.Top.Style = settingCell.Style.Border.Top.Style;
            newCell.Style.Border.Bottom.Style = settingCell.Style.Border.Bottom.Style;
            newCell.Style.ShrinkToFit = settingCell.Style.ShrinkToFit;//单元格自动适应大小
                                                                    //workSheet.Row(row + i).Height = workSheet.Row(row).Height;//设置行高
                                                                    //workSheet.Row(row + i).CustomHeight = workSheet.Row(row).CustomHeight;//自动调整行高
        }

        /// <summary>
        /// 判断单元格是否在区域内
        /// </summary>
        /// <param name="rangeLeftTopAddress"></param>
        /// <param name="rangeRightBottomAddress"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public bool IsCellInRange(string rangeLeftTopAddress, string rangeRightBottomAddress, ExcelRangeBase cell)
        {
            bool isCellInRange = false;
            if (rangeLeftTopAddress.Length < 2 || rangeRightBottomAddress.Length < 2)
            {
                return isCellInRange;
            }
            var rangeLeftTopCol = 0;
            var rangeLeftTopRow = 0;
            var rangeRightBottomCol = 0;
            var rangeRightBottomRow = 0;
            var cellCol = 0;
            var cellRow = 0;

            CalcRowCol(rangeLeftTopAddress, out rangeLeftTopRow, out rangeLeftTopCol);
            CalcRowCol(rangeRightBottomAddress, out rangeRightBottomRow, out rangeRightBottomCol);
            CalcRowCol(cell.Address, out cellRow, out cellCol);
            if (cellRow >= rangeLeftTopRow && cellRow <= rangeRightBottomRow && cellCol >= rangeLeftTopCol && cellCol <= rangeRightBottomCol)
            {
                isCellInRange = true;
            }

            return isCellInRange;
        }

        /// <summary>
        /// 判断单元格是否在区域内
        /// </summary>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        /// <param name="toRow"></param>
        /// <param name="toCol"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public bool IsCellInRange(int fromRow,int fromCol,int toRow,int toCol,ExcelRangeBase cell)
        {
            bool isCellInRange = false;

            var cellCol = 0;
            var cellRow = 0;

            CalcRowCol(cell.Address, out cellRow, out cellCol);
            if (cellRow >= fromRow && cellRow <= toRow && cellCol >= fromCol && cellCol <= toCol)
            {
                isCellInRange = true;
            }

            return isCellInRange;
        }

        /// <summary>
        /// 导出为byte[]数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public byte[] ExportToBytes(DataObject dataSource,string templatePath, string sheetName = "Sheet1")
        {
            byte[] excelBuffer = null;
            if (!File.Exists(templatePath))
            {
                return excelBuffer;
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(templatePath)))
            {
                //package.Workbook.Worksheets
                if (package.Workbook.Worksheets.Count < 1)
                {
                    return excelBuffer;
                }
                ExcelPackage newPackage = new ExcelPackage();
                newPackage.Workbook.Worksheets.Add(sheetName, package.Workbook.Worksheets[_TemplateSheetName]);
                var ws = newPackage.Workbook.Worksheets.First();
                ws = FillSheetData(ws, dataSource);
                
                excelBuffer = newPackage.GetAsByteArray();
            }
            return excelBuffer;
        }

        /// <summary>
        /// 导出为byte[]数据
        /// </summary>
        /// <param name="dataSource"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public byte[] ExportToBytes(object dataSource, string templatePath, string sheetName = "Sheet1")
        {
            return ExportToBytes(new DataObject(dataSource), templatePath, sheetName);
        }

        /// <summary>
        /// 导出为byte[]数据
        /// </summary>
        /// <param name="dataSource"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public byte[] ExportToBytes(DataSet dataSource,string templatePath,string sheetName = "Sheet1")
        {
            return ExportToBytes(new DataObject(dataSource), sheetName);
        }

        /// <summary>
        /// 字母列转换为数字列
        /// </summary>
        /// <param name="cellAddress"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        private void CalcRowCol(string cellAddress, out int row, out int col)
        {
            string rowStr = "";
            string colStr = "";
            foreach (var item in cellAddress)
            {
                if (char.IsDigit(item))
                {
                    rowStr += item;
                }
                else
                {
                    colStr += item;
                }
            }
            int.TryParse(rowStr, out row);

            col = 1;
            if (Regex.IsMatch(colStr.ToUpper(), @"[A-Z]+"))
            {
                int index = 0;
                char[] chars = colStr.ToUpper().ToCharArray();
                for (int i = 0; i < chars.Length; i++)
                {
                    index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
                }
                //col =  index - 1;
                col = index;
            }
        }
    }
}
