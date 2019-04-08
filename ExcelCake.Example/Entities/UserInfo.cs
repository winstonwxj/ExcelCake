using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelCake.Intrusive;
using System.Drawing;

namespace ExcelCake.Example
{
    [ExportEntity(EnumColor.LightGray, "用户信息")]
    [ImportEntity(titleRowIndex:1,headRowIndex:2,dataRowIndex:4)]
    public class UserInfo: ExcelBase
    {
        [Export(name:"编号", index:1,prefix:"ID:")]
        [Import(name:"编号",prefix:"ID:")]
        public int ID { set; get; }

        [Export("姓名", 2)]
        [Import("姓名")]
        public string Name { set; get; }

        [Export("性别", 3)]
        [Import("性别")]
        public string Sex { set; get; }

        [Export(name:"年龄", index:4,suffix:"岁")]
        [Import(name:"年龄",suffix:"岁",dataVerReg: @"^[1-9]\d*$", isRegFailThrowException:false)]
        public int Age { set; get; }

        [ExportMerge("联系方式")]
        [Export("电子邮件", 5)]
        [Import("电子邮件")]
        public string Email { set; get; }

        [ExportMerge("联系方式")]
        [Export("手机", 6)]
        [Import("手机")]
        public string TelPhone { set; get; }

        public override string ToString()
        {
            return string.Format($"ID:{ID},Name:{Name},Sex:{Sex},Age:{Age},Email:{Email},TelPhone:{TelPhone}");
        }
    }
}
