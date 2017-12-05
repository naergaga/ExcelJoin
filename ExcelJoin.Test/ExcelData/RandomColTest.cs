using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using ExcelJoin.Providers.Epplus;
using System.Linq;
using ExcelJoin.Models;

namespace ExcelJoin.Test.ExcelData
{
    [TestClass]
    public class RandomColTest
    {
        Random random = new Random();

        [TestMethod]
        public void ExportRandomColBook()
        {
            var file1 = new FileInfo("./files/xlsx/查岗12_05.xlsx");
            var outFilePath = "./files/xlsx/查岗_test.xlsx";

            if (File.Exists(outFilePath)) { File.Delete(outFilePath); }
            var outFile = new FileInfo(outFilePath);

            int targetCol = 1;
            int targetSheetIndex = 2;
            Sheet sheet = null;
            bool headTitle = true;

            string colName = null;
            List<string> dataList = new List<string>();
            using (ExcelPackage package = new ExcelPackage(file1))
            {
                var worksheet = package.Workbook.Worksheets[targetSheetIndex];

                var sp = new SheetProvider(worksheet, headTitle);
                sheet = sp.Get(targetCol);
                dataList = sheet.Rows.Select(t => t.Identity).ToList();
                colName = sheet.Columns[targetCol].Name;
            }

            using (ExcelPackage package = new ExcelPackage(outFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                int length = dataList.Count,
                    rowIndex = 1;

                if (headTitle)
                {
                    worksheet.Cells[rowIndex, 1].Value = sheet.Columns[targetCol-1].Name;
                    rowIndex++;
                }

                for (int i = 1; i <= length; i++,rowIndex++)
                {
                    worksheet.Cells[rowIndex, 1].Value = GetRandomFromList(dataList);
                }
                package.Save();
            }
        }

        private string GetRandomFromList(List<string> dataList)
        {
            var index = random.Next(dataList.Count);
            var value = dataList[index];
            dataList.RemoveAt(index);
            return value;
        }
    }
}
