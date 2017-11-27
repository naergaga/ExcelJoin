using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelJoin.Models;
using OfficeOpenXml;

namespace ExcelJoin.Providers.Epplus
{
    public class SheetProvider : ISheetProvider
    {
        private ExcelWorksheet sheet;
        private bool headTitle = false;

        public SheetProvider(ExcelWorksheet sheet)
        {
            this.sheet = sheet;
        }

        public Sheet GetSimple()
        {
            if (headTitle) throw new NotSupportedException("未设计支持");
            return new Sheet { Name = sheet.Name };
        }
    }
}
