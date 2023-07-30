﻿using System;
using System.Collections.Generic;
using System.IO;
using ExcelCake.Intrusive;
using ExcelCake.NoIntrusive;

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
    }
}
