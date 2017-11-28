using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelJoin.Providers.Epplus;
using System.IO;

namespace ExcelJoin.Test.Provider
{
    [TestClass]
    public class BookProviderTest
    {
        [TestMethod]
        public void GetSimple()
        {
            var path1 = "./files/xlsx/class1.xlsx";
            var file = new FileInfo(path1);
            if (!file.Exists)
            {
                Console.WriteLine("文件不存在");
            }
            Workbook workbook = new Workbook(file);
            var bp = new BookProvider(workbook.Book,true);
            var book = bp.GetSimple();
            Console.WriteLine(book);
        }
    }
}
