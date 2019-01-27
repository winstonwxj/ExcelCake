using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelCake
{
    public static class ExcelCommon
    {
        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="settingCell"></param>
        /// <param name="newCell"></param>
        public static void HandleCellStyle(ExcelRangeBase settingCell, ExcelRangeBase newCell)
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
        public static bool IsCellInRange(string rangeLeftTopAddress, string rangeRightBottomAddress, ExcelRangeBase cell)
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
        public static bool IsCellInRange(int fromRow, int fromCol, int toRow, int toCol, ExcelRangeBase cell)
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
        /// 字母列转换为数字列
        /// </summary>
        /// <param name="cellAddress"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public static void CalcRowCol(string cellAddress, out int row, out int col)
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
