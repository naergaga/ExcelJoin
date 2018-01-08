using ExcelJoin.Providers;
using ExcelJoin.Providers.EDR;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Test.Provider
{
    [TestClass]
    public class EDRBookProviderTest
    {
        [TestMethod]
        public void Get()
        {
            var bp = new EDRBookProvider();
            var path1 = "./files/xlsx/class1.xlsx";
            var book = bp.GetSimple(path1);
            Console.WriteLine(InfoProvider.GetBook(book));
        }

        [TestMethod]
        public void GetSheet()
        {
            var bp = new EDRSheetProvider();
            var path1 = "./files/xlsx/class1.xlsx";
            var book = bp.Get(path1,1,1);
            Console.WriteLine(InfoProvider.GetSheet(book));
        }
    }
}
