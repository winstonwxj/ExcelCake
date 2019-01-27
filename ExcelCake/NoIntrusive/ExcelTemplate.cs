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
        public ExcelWorksheet FillSheetData(ExcelWorksheet workSheet,ExcelObject dataSource)
        {
            if (workSheet == null || dataSource == null)
            {
                return workSheet;
            }

            #region 分析配置
            TemplateSettingSheet sheetSetting = new TemplateSettingSheet(workSheet);
            #endregion

            #region 填充数据
            foreach (var item in sheetSetting.FreeSettingList)
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

                foreach(var field in sheetSetting.FieldSettingList)
                {
                    if (ExcelCommon.IsCellInRange(item.FromRow,item.FromCol,item.ToRow,item.ToCol,field.CurrentCell))
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
            Dictionary<TemplateSettingRange, int> regionAddDic = new Dictionary<TemplateSettingRange, int>();

            foreach (var item in sheetSetting.GridSettingList)
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
                
                foreach (var field in sheetSetting.FieldSettingList)
                {
                    if (ExcelCommon.IsCellInRange(item.FromRow, item.FromCol, item.ToRow, item.ToCol, field.CurrentCell))
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
                                    var isMerge = ExcelCommon.IsCellInRange(arryMerge[0], arryMerge[1], field.CurrentCell);
                                    if (isMerge)
                                    {
                                        isFieldMerge = true;
                                        ExcelCommon.CalcRowCol(arryMerge[0], out fromRow, out fromCol);
                                        ExcelCommon.CalcRowCol(arryMerge[1], out toRow, out toCol);
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

                                    ExcelCommon.HandleCellStyle(oldRange, newRange);
                                }
                                else
                                {
                                    var oldRange = field.CurrentCell.Offset(offsetCount, 0);
                                    var newRange = field.CurrentCell.Offset(i + offsetCount, 0);

                                    ExcelCommon.HandleCellStyle(oldRange, newRange);
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
                                        ExcelCommon.HandleCellStyle(oldRange, newRange);
                                    }
                                    else
                                    {
                                        var oldRange = field.CurrentCell.Offset(offsetCount, 0);
                                        var newRange = field.CurrentCell.Offset(i + offsetCount+ fieldDatas.Count, 0);
                                        ExcelCommon.HandleCellStyle(oldRange, newRange);
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
        /// 导出为byte[]数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public byte[] ExportToBytes(ExcelObject dataSource,string templatePath, string sheetName = "Sheet1")
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
            return ExportToBytes(new ExcelObject(dataSource), templatePath, sheetName);
        }

        /// <summary>
        /// 导出为byte[]数据
        /// </summary>
        /// <param name="dataSource"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public byte[] ExportToBytes(DataSet dataSource,string templatePath,string sheetName = "Sheet1")
        {
            return ExportToBytes(new ExcelObject(dataSource), sheetName);
        }
    }
}
