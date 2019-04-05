using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    public static class ImportExtension
    {
        private static readonly string _DateTimeFormatStr = "yyyyMMddHHmmssfff";

        public static void ImportToList<T>(this IEnumerable<T> list, string cacheDirectory, string sheetName = "Sheet1", string excelName = "") where T : ExcelBase, new()
        {
            
        }

        public static void ImportAppendToList<T>(this IEnumerable<T> list, string cacheDirectory, string sheetName = "Sheet1", string excelName = "") where T : ExcelBase, new()
        {

        }
    }
}
