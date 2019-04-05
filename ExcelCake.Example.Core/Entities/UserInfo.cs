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
}
