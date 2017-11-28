using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelJoin.Models;
using OfficeOpenXml;

namespace ExcelJoin.Providers.Epplus
{
    public class BookProvider : IBookProvider
    {
        private ExcelWorkbook workbook;
        private bool headTitle;

        public BookProvider(ExcelWorkbook workbook,bool headTitle)
        {
            this.workbook = workbook;
            this.headTitle = headTitle;
        }

        public Book GetSimple()
        {
            Book book = new Book();
            book.Sheets = new List<Sheet>();
            foreach (var sheet in workbook.Worksheets)
            {
                var sp = new SheetProvider(sheet,headTitle);
                book.Sheets.Add(sp.GetSimple());
            }
            return book;
        }
    }
}
