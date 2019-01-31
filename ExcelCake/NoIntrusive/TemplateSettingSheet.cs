using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    internal class TemplateSettingSheet
    {
        private List<TemplateSettingRange> _FreeSettingList;
        private List<TemplateSettingRange> _GridSettingList;
        private List<TemplateSettingRange> _FieldSettingList;

        public List<TemplateSettingRange> FreeSettingList
        {
            get
            {
                return _FreeSettingList;
            }
        }

        public List<TemplateSettingRange> GridSettingList
        {
            get
            {
                return _GridSettingList;
            }
        }

        public List<TemplateSettingRange> FieldSettingList
        {
            get
            {
                return _FieldSettingList;
            }
        }

        private TemplateSettingSheet()
        {
            
        }

        public TemplateSettingSheet(ExcelWorksheet sheet)
        {
            _FreeSettingList = new List<TemplateSettingRange>();
            _GridSettingList = new List<TemplateSettingRange>();
            _FieldSettingList = new List<TemplateSettingRange>();

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
                        var setting = new TemplateSettingRange();
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
                            ExcelCommon.CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                            ExcelCommon.CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                            setting.FromRow = fromRow;
                            setting.FromCol = fromCol;
                            setting.ToRow = toRow;
                            setting.ToCol = toCol;
                            _GridSettingList.Add(setting);
                        }
                        else if (setting.Type == "FREE")
                        {
                            ExcelCommon.CalcRowCol(setting.AddressLeftTop, out int fromRow, out int fromCol);
                            ExcelCommon.CalcRowCol(setting.AddressRightBottom, out int toRow, out int toCol);
                            setting.FromRow = fromRow;
                            setting.FromCol = fromCol;
                            setting.ToRow = toRow;
                            setting.ToCol = toCol;
                            _FreeSettingList.Add(setting);
                        }
                        else if (setting.Type == "VALUE")
                        {
                            _FieldSettingList.Add(setting);
                        }

                    }
                    var cellValueStr = cell.Value?.ToString() ?? "";
                    cell.Value = cellValueStr.Replace("{" + item + "}", "");
                }
            }
        }
    }
}
