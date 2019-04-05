# ExcelCake

基于 EPPlus 开发的 Excel(支持 excel 2007 及之后版本)通用导入导出类库(支持.net core)。

[![](https://img.shields.io/badge/license-LGPL%20v3-blue.svg)](./LICENSE)
[![](https://img.shields.io/badge/EPPlus-v4.5.3-brightgreen.svg)](https://github.com/JanKallman/EPPlus)
![](https://img.shields.io/badge/%E4%BD%9C%E8%80%85-winstonwxj-orange.svg)

## 特性

- 快速实现 Excel 导入导出功能
- 老旧项目(.net 3.5、.net 4.0)支持
- 跨平台支持(.net standard 2.0)
- 两套实现机制:1.基于特性实现的导入导出(侵入式) 2.基于模板实现的导入导出(非侵入式)

## ExcelCake 源码结构

```
|--ExcelCake                 源码
|----Intrusive             侵入式部分
|----NoIntrusive           非侵入式部分
|--ExcelCake.Example         .net framework版Demo
|--ExcelCake.Example.Core    .net core版Demo
|--ExcelCake.Test            .net core版单元测试
```

## **作者相关**

github: https://github.com/winstonwxj
gitee: https://gitee.com/winstonwxj

## 路线图

- 基于特性表格导出(已实现)
- 基于模板表格导出(已实现)
- 基于特性表格导入(已实现)
- 基于模板表格导入
- 基于特性图表导出
- 基于模板图表导出

## 使用

## 示例

#### 基于特性实体类定义

```C#
    [ExportEntity(EnumColor.Brown,"用户信息")]
    [ImportEntity(titleRowIndex:1,headRowIndex:2,dataRowIndex:3)]
    public class UserInfo: ExcelBase
    {
        [Export("编号", 1)]
        [Import("编号")]
        public int ID { set; get; }

        [Export("姓名", 2)]
        [Import("姓名")]
        public string Name { set; get; }

        [Export("性别", 3)]
        [Import("性别")]
        public string Sex { set; get; }

        [Export("年龄", 4)]
        [Import("年龄")]
        public int Age { set; get; }

        [Export("电子邮件", 5)]
        [Import("电子邮件")]
        public string Email { set; get; }

        public string TelPhone { set; get; }

        public override string ToString()
        {
            return string.Format($"ID:{ID},Name:{Name},Sex:{Sex},Age:{Age},Email:{Email},TelPhone:{TelPhone}");
        }
    }

    [ExportEntity("账号信息")]
    public class AccountInfo:ExcelBase
    {
        [Export("编号", 1)]
        public int ID { set; get; }

        [Export("昵称", 2)]
        public string Nickname { set; get; }

        [Export("密码", 3)]
        public string Password { set; get; }

        [Export("旧密码", 4)]
        public string OldPassword { set; get; }

        [Export("状态", 5)]
        public int AccountStatus { set; get; }
    }
```

#### 基于特性导出
```C#
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
                Email = "test" + (i + 1) + "@163.com"
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
```
![avatar](./pics/pic1.PNG)

#### 基于特性多sheet多类型导出
```C#
    private static void IntrusiveMultiSheetExport()
    {
        Dictionary<string, IEnumerable<ExcelBase>> excelSheets = new Dictionary<string, IEnumerable<ExcelBase>>();

        List<UserInfo> list = new List<UserInfo>();
        List<AccountInfo> list2 = new List<AccountInfo>();
        string[] sex = new string[] { "男", "女" };

        Random random = new Random();
        for (var i = 0; i < 10000; i++)
        {
            list.Add(new UserInfo()
            {
                ID = i + 1,
                Name = "Test" + (i + 1),
                Sex = sex[random.Next(2)],
                Age = random.Next(20, 50),
                Email = "testafsdgfashgawefqwefasdfwefqwefasdggfaw" + (i + 1) + "@163.com"
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
```
![avatar](./pics/pic3.PNG)
![avatar](./pics/pic4.PNG)

#### 基于特性导入
```C#
    private static void IntrusiveImport()
    {
        var list = ExcelHelper.GetList<UserInfo>(@"C:\Users\winstonwxj\Desktop\导入文件测试.xlsx");
        foreach(var item in list)
        {
            Console.WriteLine(item);
        }
        Console.WriteLine("导入完成!");
    }
```
![avatar](./pics/pic5.PNG)

#### 模板绘制
![avatar](./pics/pic6.PNG)

#### 基于模板导出
```C#
    private static void NoIntrusiveExport()
    {
        var reportInfo = new GradeReportInfo();
        var exportTitle = "2018学年期中考试各班成绩汇总";
        reportInfo.ReportTitle = exportTitle;
        var templateFileName = "复杂格式测试模板.xlsx";

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
        var byteInfo = customTemplate.ExportToBytes(reportInfo, "Template/复杂格式测试模板.xlsx");
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
```
![avatar](./pics/pic2.PNG)

## 文档
