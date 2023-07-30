using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal class TemplateSettingRangeFree : TemplateSettingRange
    {
        private TemplateSettingRangeFree()
        {

        }

        internal static TemplateSettingRangeFree Create(ExcelRangeBase cell)
        {
            var entity = new TemplateSettingRangeFree();
            entity.Fields = new List<TemplateSettingField>();
            entity.Content = cell.Value?.ToString() ?? "";
            entity.CurrentCell = cell;
            entity.AnalyseSetting();
            return entity;
        }

        protected override void AnalyseSetting()
        {
            var cellValue = this.Content;
            var arry = cellValue.Split(new char[2] { '{', '}' }, StringSplitOptions.RemoveEmptyEntries);
            if (arry.Length == 0)
            {
                return;
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
                            case "DATASOURCE": { this.DataSource = value.ToUpper(); } break;
                            case "ADDRESSLEFTTOP": { this.AddressLeftTop = value.ToUpper(); } break;
                            case "ADDRESSRIGHTBOTTOM": { this.AddressRightBottom = value.ToUpper(); } break;
                        }
                    }
                }
                //var cellValueStr = cell.Value?.ToString() ?? "";
                //cell.Value = cellValueStr.Replace("{" + item + "}", "");
            }

            ExcelCommon.CalcRowCol(this.AddressLeftTop, out int fromRow, out int fromCol);
            ExcelCommon.CalcRowCol(this.AddressRightBottom, out int toRow, out int toCol);
            this.FromRow = fromRow;
            this.FromCol = fromCol;
            this.ToRow = toRow;
            this.ToCol = toCol;
            this.ScopeRange = this.CurrentCell.Worksheet.Cells[this.FromRow, this.FromCol, this.ToRow, this.ToCol];
        }

        internal override void Draw(ExcelWorksheet workSheet,ExcelObject dataSource)
        {
            if (string.IsNullOrEmpty(this.AddressLeftTop) || string.IsNullOrEmpty(this.AddressRightBottom))
            {
                return;
            }
            ExcelRange range = workSheet.Cells[this.AddressLeftTop + "," + this.AddressRightBottom];

            if (!dataSource.DataEntity.Keys.Contains(this.DataSource))
            {
                return;
            }
            var data = dataSource.DataEntity[this.DataSource];

            foreach (var field in Fields)
            {
                var fieldName = field.Field;
                var value = field.CurrentCell.Value?.ToString() ?? "";
                value = value.Replace(field.Content, "");
                object fieldData = null;
                if (data.Keys.Contains(field.Field))
                {
                    fieldData = data[field.Field];
                }
                field.CurrentCell.Value = value + fieldData;
            }
        }
    }
}
