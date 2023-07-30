using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal class TemplateSettingRangeGrid : TemplateSettingRange
    {
        public List<TemplateSettingRangeGrid> UnderGrid { get; set; }

        private TemplateSettingRangeGrid()
        {

        }

        internal static TemplateSettingRangeGrid Create(ExcelRangeBase cell)
        {
            var entity = new TemplateSettingRangeGrid();
            entity.Fields = new List<TemplateSettingField>();
            entity.UnderGrid = new List<TemplateSettingRangeGrid>();
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

                ExcelCommon.CalcRowCol(this.AddressLeftTop, out int fromRow, out int fromCol);
                ExcelCommon.CalcRowCol(this.AddressRightBottom, out int toRow, out int toCol);
                this.FromRow = fromRow;
                this.FromCol = fromCol;
                this.ToRow = toRow;
                this.ToCol = toCol;
                this.ScopeRange = this.CurrentCell.Worksheet.Cells[this.FromRow, this.FromCol, this.ToRow, this.ToCol];
            }
        }

        internal override void Draw(ExcelWorksheet workSheet, ExcelObject dataSource)
        {
            throw new NotImplementedException();
        }
    }
}
