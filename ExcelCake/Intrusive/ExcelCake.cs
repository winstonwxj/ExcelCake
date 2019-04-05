using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelCake.Intrusive
{
    public class ExcelHelper
    {
        private static readonly string _DateTimeFormatStr = "yyyyMMddHHmmssfff";

        public static IEnumerable<T> GetList<T>(FileInfo file, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            IEnumerable<T> list = new List<T>();

            list.ImportToList(file, importSheets, noImportSheets, importSheetsRegex, noImportSheetsRegex);

            return list;
        }

        public static IEnumerable<T> GetList<T>(string filePath, List<string> importSheets = null, List<string> noImportSheets = null, string importSheetsRegex = "", string noImportSheetsRegex = "") where T : ExcelBase, new()
        {
            return GetList<T>(new FileInfo(filePath), importSheets, noImportSheets, importSheetsRegex, noImportSheetsRegex);
        }
    }
}
