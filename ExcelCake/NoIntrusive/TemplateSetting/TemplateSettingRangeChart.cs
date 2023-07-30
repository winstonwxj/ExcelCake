using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal class TemplateSettingRangeChart: TemplateSettingRange
    {
        public string Type { set; get; }

        //public EnumChartType ChartType { get; set; }

        /// <summary>
        /// 是否自定义宽、高,默认false。为true时使用ChartWidth、ChartHeight实现图表大小
        /// </summary>
        public bool IsCustomSize { get; set; }

        public int ChartWidth { get; set; }

        public int ChartHeight { get; set; }

        public string Title { get; set; }


        private TemplateSettingRangeChart()
        {

        }

        internal static TemplateSettingRangeChart Create(ExcelRangeBase cell)
        {
            var entity = new TemplateSettingRangeChart();
            entity.Content = cell.Value?.ToString() ?? "";
            entity.CurrentCell = cell;
            entity.AnalyseSetting();
            return entity;
        }

        protected override void AnalyseSetting()
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

                    foreach (var settingItem in settingItemArry)
                    {
                        if (settingItem.ToUpper().IndexOf("TYPE") < 0)
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
                            case "FREE":
                                {
                                    var freeItem = TemplateSettingRangeFree.Create(cell);
                                    _FreeSettingList.Add(freeItem);
                                }
                                break;
                            case "GRID":
                                {
                                    var gridItem = TemplateSettingRangeGrid.Create(cell);
                                    _GridSettingList.Add(gridItem);
                                }
                                break;
                            case "CHART":
                                {
                                    var chartItem = TemplateSettingRangeChart.Create(cell);
                                    _ChartSettingList.Add(chartItem);
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


                    #region free配置分解

                    #endregion



                    TemplateSettingRange setting = null;
                    var type = "";
                    var chartSubType = "";
                    var title = "";

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
                            case "TYPE": { type = value.ToUpper(); } break;
                            case "SUBTYPE": { chartSubType = value.ToUpper(); } break;
                            case "DATASOURCE": { setting.DataSource = value.ToUpper(); } break;
                            case "ADDRESSLEFTTOP": { setting.AddressLeftTop = value.ToUpper(); } break;
                            case "ADDRESSRIGHTBOTTOM": { setting.AddressRightBottom = value.ToUpper(); } break;
                            case "FIELD":
                                {
                                    var fieldSetting = new TemplateSettingField();
                                    fieldSetting.Field = value.ToUpper();
                                    fieldSetting.Content = "{" + item + "}";
                                    fieldSetting.CurrentCell = cell;
                                    _FieldSettingList.Add(fieldSetting);
                                }
                                break;
                                //case "TITLE": { title = value; }break;
                                //case "WIDTH": {int.TryParse(value,out int width); setting.ChartWidth = width; } break;
                                //case "HEIGHT": { int.TryParse(value, out int height); setting.ChartHeight = height; } break;
                                //case "ISCUSTOMSIZE": { setting.IsCustomSize = (value == "1") ?true : false; }break;
                        }
                    }

                    if (string.IsNullOrEmpty(type))
                    {
                        continue;
                    }
                    else if (type == "GRID")
                    {

                        setting = new TemplateSettingRangeGrid();
                        setting.CurrentCell = cell;
                        setting.Content = "{" + item + "}";
                        ExcelCommon.CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                        ExcelCommon.CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                        setting.FromRow = fromRow;
                        setting.FromCol = fromCol;
                        setting.ToRow = toRow;
                        setting.ToCol = toCol;
                        _GridSettingList.Add(setting);
                    }
                    else if (type == "FREE")
                    {
                        setting = new TemplateSettingRangeFree();
                        setting.CurrentCell = cell;
                        setting.Content = "{" + item + "}";
                        ExcelCommon.CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                        ExcelCommon.CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                        setting.FromRow = fromRow;
                        setting.FromCol = fromCol;
                        setting.ToRow = toRow;
                        setting.ToCol = toCol;
                        _FreeSettingList.Add(setting);
                    }
                    else if (type == "CHART")
                    {
                        setting = new TemplateSettingRangeChart();
                        setting.CurrentCell = cell;
                        setting.Content = "{" + item + "}";
                        ExcelCommon.CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                        ExcelCommon.CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                        setting.FromRow = fromRow;
                        setting.FromCol = fromCol;
                        setting.ToRow = toRow;
                        setting.ToCol = toCol;
                        _ChartSettingList.Add(setting);

                        //if (string.IsNullOrEmpty(chartSubType))
                        //{
                        //    setting.ChartType = EnumChartType.Chart;
                        //}
                        //else
                        //{
                        //    switch (chartSubType)
                        //    {
                        //        case "CHART": { setting.ChartType = EnumChartType.Chart; } break;
                        //        case "BARCHART": { setting.ChartType = EnumChartType.BarChart; } break;
                        //        case "BUBBLECHART": { setting.ChartType = EnumChartType.BubbleChart; } break;
                        //        case "DOUGHNUTCHART": { setting.ChartType = EnumChartType.DoughnutChart; } break;
                        //        case "LINECHART": { setting.ChartType = EnumChartType.LineChart; } break;
                        //        case "OFPIECHART": { setting.ChartType = EnumChartType.OfPieChart; } break;
                        //        case "PIECHART": { setting.ChartType = EnumChartType.PieChart; } break;
                        //        case "RADARCHART": { setting.ChartType = EnumChartType.RadarChart; } break;
                        //        case "SCATTERCHART": { setting.ChartType = EnumChartType.ScatterChart; } break;
                        //        case "SURFACECHART": { setting.ChartType = EnumChartType.SurfaceChart; } break;
                        //        default: { setting.ChartType = EnumChartType.Chart; } break;
                        //    }
                        //}
                    }
                    else if (type == "VALUE")
                    {
                        _FieldSettingList.Add(setting);
                    }

                }
                var cellValueStr = cell.Value?.ToString() ?? "";
                cell.Value = cellValueStr.Replace("{" + item + "}", "");
            }
        }

        internal override void Draw(ExcelWorksheet workSheet, ExcelObject dataSource)
        {
            throw new NotImplementedException();
        }
    }

    internal enum EnumChartType
    {
        None = 0,
        Chart = 1,
        BarChart = 2,
        BubbleChart = 3,
        DoughnutChart = 4,
        LineChart = 5,
        OfPieChart = 6,
        PieChart = 7,
        RadarChart = 8,
        ScatterChart = 9,
        SurfaceChart = 10,
    }
}
