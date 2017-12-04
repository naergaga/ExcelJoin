using ExcelJoin.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Actions
{
    public class JoinAction
    {
        public JoinAction()
        {

        }

        public void Export(Sheet sheet1, Sheet sheet2, string outpath, string sheetName, bool headTitle1 = false, bool headTitle2 = false)
        {
            var newFile = new FileInfo(outpath);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(outpath);
            }
            var query = from s1row in sheet1.Rows
                        join s2row in sheet2.Rows
                        on s1row.Identity equals s2row.Identity
                        select new { Identity = s1row.Identity, Data1 = s1row.Data, Data2 = s2row.Data };
            var list = query.ToList();

            using (var package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                int rowIndex = 1;
                //write column title
                int colIndex = 1;
                if (headTitle1)
                {
                    for (int i = 0; i < sheet1.Columns.Count; i++, colIndex++)
                    {
                        worksheet.Cells[rowIndex, colIndex].Value = sheet1.Columns[i].Name;
                    }
                }
                if (headTitle2)
                {
                    for (int i = 0; i < sheet2.Columns.Count; i++, colIndex++)
                    {
                        worksheet.Cells[rowIndex, colIndex].Value = sheet2.Columns[i].Name;
                    }
                }

                if (headTitle1 || headTitle2)
                {
                    rowIndex += 1;
                }

                //write column data
                for (int i = 0; i < list.Count; i++, rowIndex++)
                {
                    var rowData = list[i];
                    var data1 = rowData.Data1;
                    var data2 = rowData.Data2;
                    var col = 1;

                    for (int i2 = 0; i2 < data1.Count; i2++, col++)
                    {
                        worksheet.Cells[rowIndex, col].Value = data1[i2].Value;
                    }

                    for (int i2 = 0; i2 < data2.Count; i2++, col++)
                    {
                        worksheet.Cells[rowIndex, col].Value = data2[i2].Value;
                    }

                }
                package.Save();
            }
        }
    }
}
