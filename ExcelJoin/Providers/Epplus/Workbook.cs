using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Providers.Epplus
{
    public class Workbook : IDisposable
    {
        public ExcelPackage Package { get; set; }
        public ExcelWorkbook Book
        {
            get
            {
                return Package.Workbook;
            }
        }

        public Workbook(FileInfo inputFile)
        {
            Package = new ExcelPackage(inputFile);
        }

        public void Dispose()
        {
            Package.Dispose();
        }
    }
}
