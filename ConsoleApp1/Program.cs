using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            DrawTest();
            Console.WriteLine("Done!");
            Console.ReadKey();
        }

        static void DrawTest()
        {
            //ExcelBarChart
            //ExcelBubbleChart
            //ExcelChart
            //ExcelDoughnutChart
            //ExcelLineChart
            //ExcelOfPieChart
            //ExcelPieChart
            //ExcelRadarChart
            //ExcelScatterChart
            //ExcelSurfaceChart

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test.xlsx");
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            using (ExcelPackage package = new ExcelPackage())
            {
                var hideWorksheet = package.Workbook.Worksheets.Add("dic");
                hideWorksheet.Hidden = eWorkSheetHidden.VeryHidden;
                hideWorksheet.Cells.Style.WrapText = true;
                hideWorksheet.Cells[1, 1].Value = "名称";
                hideWorksheet.Cells[1, 2].Value = "价格";
                hideWorksheet.Cells[1, 3].Value = "销量";

                hideWorksheet.Cells[2, 1].Value = "大米";
                hideWorksheet.Cells[2, 2].Value = 56;
                hideWorksheet.Cells[2, 3].Value = 100;

                hideWorksheet.Cells[3, 1].Value = "玉米";
                hideWorksheet.Cells[3, 2].Value = 45;
                hideWorksheet.Cells[3, 3].Value = 150;

                hideWorksheet.Cells[4, 1].Value = "小米";
                hideWorksheet.Cells[4, 2].Value = 38;
                hideWorksheet.Cells[4, 3].Value = 130;

                hideWorksheet.Cells[5, 1].Value = "糯米";
                hideWorksheet.Cells[5, 2].Value = 22;
                hideWorksheet.Cells[5, 3].Value = 200;

                using (ExcelRange range = hideWorksheet.Cells[1, 1, 5, 3])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                using (ExcelRange range = hideWorksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.Font.Name = "微软雅黑";
                    range.Style.Font.Size = 12;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                }

                hideWorksheet.Cells[1, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[1, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                hideWorksheet.Cells[2, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[2, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[2, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                hideWorksheet.Cells[3, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[3, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[3, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                hideWorksheet.Cells[4, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[4, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[4, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                hideWorksheet.Cells[5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[5, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                hideWorksheet.Cells[5, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("test");

                worksheet.Cells.Style.WrapText = true;
                worksheet.View.ShowGridLines = false;//去掉sheet的网格线

                //ExcelChart chart = worksheet.Drawings.AddChart("chart", eChartType.ColumnClustered);
                ExcelChart chart = worksheet.Drawings.AddChart("chart", eChartType.ConeCol);

                //ExcelChartSerie serie = chart.Series.Add(worksheet.Cells[2, 3, 5, 3], worksheet.Cells[2, 1, 5, 1]);
                //serie.HeaderAddress = worksheet.Cells[1, 3];

                ExcelChartSerie serie = chart.Series.Add("dic!$C$1:$C$5", "dic!$A$1:$A$5");
                serie.HeaderAddress = hideWorksheet.Cells[1, 3];

                chart.SetPosition(150, 10);
                chart.SetSize(500, 300);
                chart.Title.Text = "销量走势";
                chart.Title.Font.Color = Color.FromArgb(89, 89, 89);
                chart.Title.Font.Size = 15;
                chart.Title.Font.Bold = true;
                chart.Style = eChartStyle.Style15;
                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);

                package.SaveAs(new FileInfo(path));
            }

        }
    }
}
