using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    internal class TemplateSettingSheet
    {
        private List<TemplateSettingRangeFree> _FreeSettingList;
        private List<TemplateSettingRangeGrid> _GridSettingList;
        //private List<TemplateSettingRangeChart> _ChartSettingList;
        private List<TemplateSettingField> _FieldSettingList;
        public List<TemplateSettingRangeFree> FreeSettingList
        {
            get
            {
                return _FreeSettingList;
            }
        }

        public List<TemplateSettingRangeGrid> GridSettingList
        {
            get
            {
                return _GridSettingList;
            }
        }

        //public List<TemplateSettingRangeChart> ChartSettingList
        //{
        //    get
        //    {
        //        return _ChartSettingList;
        //    }
        //}

        private TemplateSettingSheet()
        {
            
        }

        public TemplateSettingSheet(ExcelWorksheet sheet)
        {
            _FreeSettingList = new List<TemplateSettingRangeFree>();
            _GridSettingList = new List<TemplateSettingRangeGrid>();
            //_ChartSettingList = new List<TemplateSettingRangeChart>();
            _FieldSettingList = new List<TemplateSettingField>();

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

                        foreach(var settingItem in settingItemArry)
                        {
                            if (settingItem.ToUpper().IndexOf("TYPE") <0 )
                            {
                                continue;
                            }

                            var arrItem = settingItem.Split(new char[1] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                            if (arrItem.Length < 2)
                            {
                                continue;
                            }

                            var key = arrItem[0];
                            var value = arrItem[1];

                            switch (value.ToUpper())
                            {
                                case "FREE": {
                                        var freeItem = TemplateSettingRangeFree.Create(cell);
                                        _FreeSettingList.Add(freeItem);
                                    } break;
                                case "GRID":
                                    {
                                        var gridItem =TemplateSettingRangeGrid.Create(cell);
                                        _GridSettingList.Add(gridItem);
                                    }
                                    break;
                                case "CHART":
                                    {
                                        //var chartItem = TemplateSettingRangeChart.Create(cell);
                                        //_ChartSettingList.Add(chartItem);
                                    }
                                    break;
                                case "VALUE":
                                    {
                                        var fieldItem = TemplateSettingField.Create(cell);
                                        _FieldSettingList.Add(fieldItem);
                                    }
                                    break;
                            }
                        }
                    }
                    var cellValueStr = cell.Value?.ToString() ?? "";
                    cell.Value = cellValueStr.Replace("{" + item + "}", "");
                }

                //自由格式
                foreach(var free in _FreeSettingList)
                {
                    foreach(var field in _FieldSettingList)
                    {
                        var isContain = ExcelCommon.IsCellInRange(free.AddressLeftTop, free.AddressRightBottom, field.CurrentCell);

                        if (isContain)
                        {
                            free.Fields.Add(field);
                        }
                    }
                }

                //表格
                foreach(var grid in _GridSettingList)
                {
                    foreach (var field in _FieldSettingList)
                    {
                        var isContain = ExcelCommon.IsCellInRange(grid.AddressLeftTop, grid.AddressRightBottom, field.CurrentCell);

                        if (isContain)
                        {
                            grid.Fields.Add(field);
                        }
                    }
                }


                //图表
            }
        }
    }
}
