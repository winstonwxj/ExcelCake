using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using ExcelCake.Intrusive;
using ExcelCake.NoIntrusive;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml;

namespace ExcelCake.Example.Core
{
    class Program
    {
        static void Main(string[] args)
        {
            //IntrusiveExport();
            //IntrusiveMultiSheetExport();
            NoIntrusiveExport();
            //IntrusiveImport();
            //DrawTest();
            Console.ReadKey();
        }

        private static void IntrusiveExport()
        {
            List<UserInfo> list = new List<UserInfo>();
            string[] sex = new string[] { "男", "女" };
            Random random = new Random();
            for (var i = 0; i < 100; i++)
            {
                list.Add(new UserInfo()
                {
                    ID = i + 1,
                    Name = "Test" + (i + 1),
                    Sex = sex[random.Next(2)],
                    Age = random.Next(20, 50),
                    Email = "test" + (i + 1) + "@163.com",
                    TelPhone = "1399291" + random.Next(1000,9999)
                });
            }
            var temp = list.ExportToExcelBytes(); //导出为byte[]

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Export");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var exportTitle = "导出文件";
            var filePath = Path.Combine(path, exportTitle + DateTime.Now.Ticks + ".xlsx");
            FileInfo file = new FileInfo(filePath);
            File.WriteAllBytes(file.FullName, temp);
            Console.WriteLine("IntrusiveExport导出完成!");
        }

        private static void IntrusiveMultiSheetExport()
        {
            Dictionary<string, IEnumerable<ExcelBase>> excelSheets = new Dictionary<string, IEnumerable<ExcelBase>>();

            List<UserInfo> list = new List<UserInfo>();
            List<AccountInfo> list2 = new List<AccountInfo>();
            string[] sex = new string[] { "男", "女" };

            Random random = new Random();
            for (var i = 0; i < 100; i++)
            {
                list.Add(new UserInfo()
                {
                    ID = i + 1,
                    Name = "Test" + (i + 1),
                    Sex = sex[random.Next(2)],
                    Age = random.Next(20, 50),
                    Email = "testafsdgfashgawefqwefasdfwefqwefasdggfaw" + (i + 1) + "@163.com",
                    TelPhone = "1399291" + random.Next(1000, 9999)
                });
                list2.Add(new AccountInfo()
                {
                    ID = i + 1,
                    Nickname = "nick" + (i + 1),
                    Password = random.Next(111111, 999999).ToString(),
                    OldPassword = random.Next(111111, 999999).ToString(),
                    AccountStatus = random.Next(2)
                });
            }
            excelSheets.Add("sheet1", list);
            excelSheets.Add("sheet2", list2);


            var temp = excelSheets.ExportMultiToBytes(); //导出为byte[]

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Export");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var exportTitle = "导出文件";
            var filePath = Path.Combine(path, exportTitle + DateTime.Now.Ticks + ".xlsx");
            FileInfo file = new FileInfo(filePath);
            File.WriteAllBytes(file.FullName, temp);
            Console.WriteLine("IntrusiveMultiSheetExport导出完成!");
        }

        private static void NoIntrusiveExport()
        {
            var reportInfo = new GradeReportInfo();
            var exportTitle = "2018学年期中考试各班成绩汇总";
            reportInfo.ReportTitle = exportTitle;
            //var templateFileName = "复杂格式测试模板.xlsx";
            var templateFileName = "复杂格式测试模板(新语法).xlsx";

            #region 构造数据
            var list1 = new List<ClassInfo>();
            list1.Add(new ClassInfo()
            {
                ClassName = "班级1",
                PassCountSubject1 = 20,
                PassCountSubject2 = 15,
                PassCountSubject3 = 10,
                PassCountSubject4 = 13,
                PassCountSubject5 = 25
            });
            list1.Add(new ClassInfo()
            {
                ClassName = "班级2",
                PassCountSubject1 = 19,
                PassCountSubject2 = 20,
                PassCountSubject3 = 17,
                PassCountSubject4 = 11,
                PassCountSubject5 = 19
            });
            list1.Add(new ClassInfo()
            {
                ClassName = "班级3",
                PassCountSubject1 = 17,
                PassCountSubject2 = 23,
                PassCountSubject3 = 12,
                PassCountSubject4 = 16,
                PassCountSubject5 = 21
            });
            list1.Add(new ClassInfo()
            {
                ClassName = "班级4",
                PassCountSubject1 = 23,
                PassCountSubject2 = 17,
                PassCountSubject3 = 16,
                PassCountSubject4 = 14,
                PassCountSubject5 = 22
            });
            list1.Add(new ClassInfo()
            {
                ClassName = "班级5",
                PassCountSubject1 = 23,
                PassCountSubject2 = 17,
                PassCountSubject3 = 16,
                PassCountSubject4 = 14,
                PassCountSubject5 = 22
            });
            var list2 = new List<ClassInfo>();
            list2.Add(new ClassInfo()
            {
                ClassName = "班级1",
                ScoreAvgSubject1 = 81.25,
                ScoreAvgSubject2 = 65.75,
                ScoreAvgSubject3 = 79.05,
                ScoreAvgSubject4 = 59.15,
                ScoreAvgSubject5 = 83.05
            });
            list2.Add(new ClassInfo()
            {
                ClassName = "班级2",
                ScoreAvgSubject1 = 79.25,
                ScoreAvgSubject2 = 63.75,
                ScoreAvgSubject3 = 71.05,
                ScoreAvgSubject4 = 62.15,
                ScoreAvgSubject5 = 85
            });
            list2.Add(new ClassInfo()
            {
                ClassName = "班级3",
                ScoreAvgSubject1 = 71.5,
                ScoreAvgSubject2 = 63.25,
                ScoreAvgSubject3 = 75.25,
                ScoreAvgSubject4 = 61.25,
                ScoreAvgSubject5 = 80.05
            });
            list2.Add(new ClassInfo()
            {
                ClassName = "班级4",
                ScoreAvgSubject1 = 84.5,
                ScoreAvgSubject2 = 61.25,
                ScoreAvgSubject3 = 75.25,
                ScoreAvgSubject4 = 57.35,
                ScoreAvgSubject5 = 81.5
            });
            var list3 = new List<ClassInfo>();
            list3.Add(new ClassInfo()
            {
                ClassName = "班级1",
                ScoreTotalMax = 432,
                ScoreTotalAvg = 315.25,
                ScoreTotalPassRate = 47.25
            });
            list3.Add(new ClassInfo()
            {
                ClassName = "班级2",
                ScoreTotalMax = 466.5,
                ScoreTotalAvg = 330.75,
                ScoreTotalPassRate = 44.75
            });
            list3.Add(new ClassInfo()
            {
                ClassName = "班级3",
                ScoreTotalMax = 422,
                ScoreTotalAvg = 345.25,
                ScoreTotalPassRate = 51.05
            });
            list3.Add(new ClassInfo()
            {
                ClassName = "班级4",
                ScoreTotalMax = 444,
                ScoreTotalAvg = 335.25,
                ScoreTotalPassRate = 46.15
            });
            #endregion

            reportInfo.List1 = list1;
            reportInfo.List2 = list2;
            reportInfo.List3 = list3;

            ExcelTemplate customTemplate = new ExcelTemplate(templateFileName);
            var byteInfo = customTemplate.ExportToBytes(reportInfo, "Template/"+ templateFileName);
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Export");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var filePath = Path.Combine(path, exportTitle + DateTime.Now.Ticks + ".xlsx");
            FileInfo file = new FileInfo(filePath);
            File.WriteAllBytes(file.FullName, byteInfo);
            Console.WriteLine("NoIntrusiveExport导出完成!");
        }

        private static void IntrusiveImport()
        {
            var list = ExcelHelper.GetList<UserInfo>(@"C:\Users\winstonwxj\Desktop\导入文件测试.xlsx");
            foreach(var item in list)
            {
                Console.WriteLine(item);
            }
            Console.WriteLine("导入完成!");
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

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xlsx");
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
