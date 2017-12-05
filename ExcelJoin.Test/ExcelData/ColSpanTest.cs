using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace ExcelJoin.Test.ExcelData
{
    [TestClass]
    public class ColSpanTest
    {
        [TestMethod]
        public void TestColSpan()
        {
            var file1 = new FileInfo("./files/xlsx/class1.xlsx");
            using (ExcelPackage package = new ExcelPackage(file1))
            {
                var worksheet = package.Workbook.Worksheets[3];
                var rowLength = worksheet.Dimension.Rows;
                var ColLength = worksheet.Dimension.Columns;
                for (int ri = 1; ri <= rowLength; ri++)
                {
                    for (int ci = 1; ci <= ColLength; ci++)
                    {
                        Console.Write($"{ri}-{ci} {worksheet.GetValue(ri, ci)}  ");
                    }
                    Console.WriteLine();
                }
            }
        }
    }
}
